'   Main.EventReceiver.vb
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

Option Explicit On
Option Strict On

Imports System
Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports Microsoft.SharePoint
Imports Microsoft.SharePoint.Security
Imports Microsoft.SharePoint.Administration

''' <summary>
''' This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
''' </summary>
''' <remarks>
''' The GUID attached to this class may be used during packaging and should not be modified.
''' </remarks>

<GuidAttribute("a885d247-f5e8-4456-abd2-6cfebb2bdfde")> _
Public Class MainEventReceiver
    Inherits SPFeatureReceiver

    Public Sub RunMetaDataNavExpansionJob(Installing As Boolean, properties As SPFeatureReceiverProperties)
        Dim JobName As String = "MetaDataNavExpansionJob"

        'Ensure job doesn't already exist (delete if it does)
        Dim query = From job As SPJobDefinition In properties.Definition.Farm.TimerService.JobDefinitions Where job.Name.Equals(JobName) Select job
        Dim myJobDefinition As SPJobDefinition = query.FirstOrDefault()
        If myJobDefinition IsNot Nothing Then myJobDefinition.Delete()

        Dim myJob As New MetaDataNavExpansionJob(JobName, SPFarm.Local.TimerService, Installing)

        'Get that job going!
        myJob.Title = String.Format("Configuring MetaData Navigation for {0} Expansion", IIf(Installing, "First Level", "Default"))
        myJob.Update()
        myJob.RunNow()
    End Sub


    Public Overrides Sub FeatureActivated(ByVal properties As SPFeatureReceiverProperties)
        RunMetaDataNavExpansionJob(True, properties)
    End Sub

    Public Overrides Sub FeatureDeactivating(ByVal properties As SPFeatureReceiverProperties)
        RunMetaDataNavExpansionJob(False, properties)
    End Sub

End Class
