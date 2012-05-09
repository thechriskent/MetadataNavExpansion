'   MetaDataNavExpansionJob.vb
'   WireBear MetaDataNavExpansion
'
'   Created by Chris Kent
'   Copyright 2012 WireBear. All rights reserved
'
#Region "License Information"
'   WireBear MetaDataNavExpansion is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'
'   WireBear MetaDataNavExpansion is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with WireBear MetaDataNavExpansion - License.txt
'   If not, see <http://www.gnu.org/licenses/>.
#End Region

Imports Microsoft.SharePoint.Administration
Imports Microsoft.SharePoint.Utilities
Imports System.Text.RegularExpressions

Public Class MetaDataNavExpansionJob
    Inherits SPServiceJobDefinition

#Region "Properties"

    Private _userControlPath As String
    Public ReadOnly Property UserControlPath() As String
        Get
            If String.IsNullOrEmpty(_userControlPath) Then _userControlPath = SPUtility.GetGenericSetupPath("TEMPLATE\CONTROLTEMPLATES\MetadataNavTree.ascx")
            Return _userControlPath
        End Get
    End Property

    Private _userControlBackupPath As String
    Public ReadOnly Property UserControlBackupPath() As String
        Get
            If String.IsNullOrEmpty(_userControlBackupPath) Then _userControlBackupPath = SPUtility.GetGenericSetupPath("TEMPLATE\CONTROLTEMPLATES\MetadataNavTree.ascx.bak")
            Return _userControlBackupPath
        End Get
    End Property

    Private Const InstallingKey As String = "DocIconJob_InstallingKey"
    Private Property _installing() As Boolean
        Get
            If Properties.ContainsKey(InstallingKey) Then
                Return Convert.ToBoolean(Properties(InstallingKey))
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If Properties.ContainsKey(InstallingKey) Then
                Properties(InstallingKey) = value.ToString
            Else
                Properties.Add(InstallingKey, value.ToString)
            End If
        End Set
    End Property

#End Region

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(JobName As String, service As SPService, Installing As Boolean)
        MyBase.New(JobName, service)
        _installing = Installing
    End Sub

    Public Overrides Sub Execute(jobState As Microsoft.SharePoint.Administration.SPJobState)
        AdjustMetaDataNavExpansion()
    End Sub

    Private Sub AdjustMetaDataNavExpansion()
        If _installing Then
            If My.Computer.FileSystem.FileExists(UserControlPath) Then
                'Backup the original
                My.Computer.FileSystem.CopyFile(UserControlPath, UserControlBackupPath, True)
                Dim contents As String = My.Computer.FileSystem.ReadAllText(UserControlPath)

                'Replace the Expansion with First Level Expansion
                My.Computer.FileSystem.WriteAllText(UserControlPath, Regex.Replace(contents, "ExpandDepth=""\d+""", "ExpandDepth=""2"""), False)
            End If
        Else
            If My.Computer.FileSystem.FileExists(UserControlBackupPath) Then
                'Restore the original
                My.Computer.FileSystem.MoveFile(UserControlBackupPath, UserControlPath, True)
            End If
        End If
    End Sub

End Class
