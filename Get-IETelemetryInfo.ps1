<#
.Synopsis
   The Get-IETelemetryURLInfo retrieves script retrieves the Internet Explorer Telemetry information from the specified computers
.DESCRIPTION
   The Get-IETelemetryURLInfo retrieves script retrieves the Internet Explorer Telemetry information from the specified computers.
   The script also translates the ActiveX Guid when it's found in the ActiveX reference list. Furthermore the scrpt translates the
   DocMode Reason, the Browser Mode Reason and the Zone information. 

.PARAMETER Computername
    One or mutliple computer names
.EXAMPLE
   Get-IETelemetryURLInfo -Computername Client01,Client02

        ComputerName                : chr5bi01
        IESystemInfo                : \\CHR5BI01\ROOT\cimv2\IETelemetry:IESystemInfo.SystemKey="SystemKey"
        IECountInfo                 : \\CHR5BI01\ROOT\cimv2\IETelemetry:IECountInfo.CountKey="CountKey"
        ActiveXGUID                 : 
        ActiveXDetail               : {@{URL=http://www.verboon.info/; ActiveXGUID=No ActiveX detected; Description=}}
        BrowserStateReason          : 12
        BrowserStateReasonDesc      : 
        CrashCount                  : 0
        DocMode                     : 11
        DocModeReason               : 9
        DocModeReasonDesc           : Document mode is the result of the page's doctype and the browser mode
        Domain                      : verboon.info
        HangCount                   : 0
        MostRecentNavigationFailure : 
        NavigationFailureCount      : 0
        NumberOfVisits              : 1
        URL                         : http://www.verboon.info/
        Zone                        : 3
        ZoneDescription             : INTERNET
  
   This command retrieves the Internet Explorer Telemetry data from the specified computers. 

.EXAMPLE
    Get-IETelemetryURLInfo -ComputerName chr5bi01 -ActiveX

    URL                                      ActiveXGUID                              Description                             
    ---                                      -----------                              -----------                             
    http://www.foosample1.com/               No ActiveX detected                                                              
    http://www.verboon.info/                 No ActiveX detected                                                              
    http://www.foosample2.com/               No ActiveX detected                                                              
    http://www.foosample3.com/               {D27CDB6E-AE6D-11CF-96B8-444553540000}   Shockwave Flash Object                  
    http://intranet.foocorp.com/             {F6D90F16-9C73-11D3-B32E-00C04F990BB4}   XML HTTP                                
    http://intranet.foocorp.com/             {ED8C108E-4349-11D2-91A4-00C04F7969E8}   XML HTTP Request        

    This command returns URL and ActiveX information. When the ActiveX GUID is found in the ActiveX reference list
    defined within the script the name is displayed in the Description. 
  

.NOTES
  To use this script clients must be configured to collect Internet Explorer usage data as described in the 
  articles referenced below. 
  http://technet.microsoft.com/en-us/library/dn833204.aspx  
  http://blogs.msdn.com/b/ie/archive/2014/10/30/making-it-easier-for-enterprises-to-stay-up-to-date.aspx

  version 1.0   29-NOV-2014, Alex Verboon

#>
function Get-IETelemetryURLInfo
{
    [CmdletBinding()]
    Param
    (
        # 
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
        [string[]]$ComputerName,
        [switch]$ActiveX
    )

    Begin
    {
        # Collection of known ActiveX     
        $activexlist = @{
        "{47833539-D0C5-4125-9FA8-0819E2EAAC93}" = "Adobe Acrobat Create PDF Toolbar"
        #"{47833539-D0C5-4125-9FA8-0819E2EAAC93}" = "Adobe PDF"
        "{28BCCB9A-E66B-463C-82A4-09F320DE94D7}" = "F12 Developer Tools"
        #"{28BCCB9A-E66B-463C-82A4-09F320DE94D7}" = "F12 Developer Tools"
        "{AE7CD045-E861-484f-8273-0445EE161910}" = "Adobe Acrobat Create PDF Helper"
        #"{AE7CD045-E861-484f-8273-0445EE161910}" = "Adobe PDF Conversion Toolbar Helper"
        "{B4F3A835-0E21-4959-BA22-42B3008E02FF}" = "Office Document Cache Handler"
        "{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}" = "Microsoft SkyDrive Pro Browser Helper"
        "{F4971EE7-DAA0-4053-9964-665D8EE6A077}" = "Adobe Acrobat Create PDF from Selection"
        #"{F4971EE7-DAA0-4053-9964-665D8EE6A077}" = "SmartSelect Class"
        "{48E73304-E1D6-4330-914C-F5F514E3486C}" = "Send to OneNote"
        "{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}" = "Lync Click to Call"
        #"{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}" = "Lync add-on"
        #"{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}" = "Lync Browser Helper"
        "{FFFDC614-B694-4AE6-AB38-5D6374584B52}" = "OneNote Linked Notes"
        "{10336656-40D7-4530-BCC0-86CD3D77D25F}" = "MeetingJoinHelper Class"
        "{1542FC7D-8D51-43D5-B757-67C763F27BF4}" = "Microsoft Lync Web App Version Plug-in (64-bit)"
        "{25336920-03F9-11CF-8FD0-00AA00686F13}" = "HTML Document"
        "{2933BF90-7B36-11D2-B20E-00C04F983E60}" = "XML DOM Document"
        "{4FCEE402-10E9-4446-AE0F-AE48D6D62E9A}" = "Groove Site Client ActiveX"
        "{52A2AAAE-085D-4187-97EA-8C30DB990436}" = "HHCtrl Object"
        #"{52A2AAAE-085D-4187-97EA-8C30DB990436}" = "HHCtrl Object"
        "{6BF52A52-394A-11D3-B153-00C04F79FAA6}" = "Windows Media Player"
        "{7ECF6F97-B4F3-4168-9835-F59C06D7875F}" = "Microsoft Lync Web App Plug-in (64-bit)"
        "{8075831E-5146-11D5-A672-00B0D022E945}" = "SharepointOpenXMLDocuments"
        "{8856F961-340A-11D0-A96B-00C04FD705A2}" = "Microsoft Web Browser"
        "{96CAE7ED-F021-4FEB-A5E9-7CC58829A67A}" = "Microsoft Lync Web App Plug-in (64-bit)"
        "{971127BB-259F-48C2-BD75-5F97A3331551}" = "Microsoft RDP Client Control (redistributable) - version 3a"
        #"{971127BB-259F-48c2-BD75-5F97A3331551}" = "Microsoft RDP Client Control (redistributable) - version 3a"
        "{D27CDB6E-AE6D-11CF-96B8-444553540000}" = "Shockwave Flash Object"
        "{DFEAF541-F3E1-4C24-ACAC-99C30715084A}" = "Microsoft Silverlight"
        #"{DFEAF541-F3E1-4c24-ACAC-99C30715084A}" = "Microsoft Silverlight"
        "{ED8C108E-4349-11D2-91A4-00C04F7969E8}" = "XML HTTP Request"
        "{EE09B103-97E0-11CF-978F-00A02463E06F}" = "Scripting.Dictionary"
        "{F5078F32-C551-11D3-89B9-0000F81FE221}" = "XML DOM Document 3.0"
        "{F6D90F11-9C73-11D3-B32E-00C04F990BB4}" = "XML DOM Document"
        "{F6D90F16-9C73-11D3-B32E-00C04F990BB4}" = "XML HTTP"
        "{54CE37E0-9834-41ae-9896-4DAB69DC022B}" = "Microsoft RDP Client Control (redistributable) - version 5a"
        "{6A6F4B83-45C5-4ca9-BDD9-0D81C12295E4}" = "Microsoft RDP Client Control (redistributable) - version 4a"
        "{41B23C28-488E-4E5C-ACE2-BB0BBABE99E8}" = "HHCtrl Object"
        "{8AD9C840-044E-11D1-B3E9-00805F499D93}" = "Java Plug-in 10.21.2"
        #"{8AD9C840-044E-11D1-B3E9-00805F499D93}" = "Java Plug-in 1.6.0_16"
        #"{8AD9C840-044E-11D1-B3E9-00805F499D93}" = "Java Plug-in 10.55.2"
        #"{8AD9C840-044E-11D1-B3E9-00805F499D93}" = "Java Plug-in 1.6.0_45"
        #"{8AD9C840-044E-11D1-B3E9-00805F499D93}" = "Java Plug-in 10.40.2"
        #"{8AD9C840-044E-11D1-B3E9-00805F499D93}" = "Java Plug-in 1.6.0_43"
        #"{8AD9C840-044E-11D1-B3E9-00805F499D93}" = "Java Plug-in 1.6.0_24"
        #"{8AD9C840-044E-11D1-B3E9-00805F499D93}" = "Java Plug-in 1.5.0_14"
        #"{8AD9C840-044E-11D1-B3E9-00805F499D93}" = "Java Plug-in 1.5.0_22"
        #"{8AD9C840-044E-11D1-B3E9-00805F499D93}" = "Java Plug-in 10.0.0"
        #"{8AD9C840-044E-11D1-B3E9-00805F499D93}" = "Java Plug-in 1.6.0_21"
        #"{8AD9C840-044E-11D1-B3E9-00805F499D93}" = "Java Plug-in 1.6.0_37"
        "{8D9563A9-8D5F-459B-87F2-BA842255CB9A}" = "Forefront UAG client components"
        "{CAFEEFAC-0016-0000-0045-ABCDEFFEDCBA}" = "Java Plug-in 1.6.0_45"
        "{CAFEEFAC-0017-0000-0021-ABCDEFFEDCBA}" = "Java Plug-in 1.7.0_21"
        "{CAFEEFAC-FFFF-FFFF-FFFF-ABCDEFFEDCBA}" = "Java Plug-in 1.7.0_21"
        #"{CAFEEFAC-FFFF-FFFF-FFFF-ABCDEFFEDCBA}" = "Java Plug-in 1.6.0_16"
        #"{CAFEEFAC-FFFF-FFFF-FFFF-ABCDEFFEDCBA}" = "Java Plug-in 10.55.2"
        #"{CAFEEFAC-FFFF-FFFF-FFFF-ABCDEFFEDCBA}" = "Java Plug-in 1.6.0_45"
        #"{CAFEEFAC-FFFF-FFFF-FFFF-ABCDEFFEDCBA}" = "Java Plug-in 10.40.2"
        #"{CAFEEFAC-FFFF-FFFF-FFFF-ABCDEFFEDCBA}" = "Java Plug-in 1.6.0_43"
        #"{CAFEEFAC-FFFF-FFFF-FFFF-ABCDEFFEDCBA}" = "Java Plug-in 1.6.0_24"
        #"{CAFEEFAC-FFFF-FFFF-FFFF-ABCDEFFEDCBA}" = "Java Plug-in 1.5.0_14"
        #"{CAFEEFAC-FFFF-FFFF-FFFF-ABCDEFFEDCBA}" = "Java Plug-in 1.5.0_22"
        #"{CAFEEFAC-FFFF-FFFF-FFFF-ABCDEFFEDCBA}" = "Java Plug-in 1.7.0"
        #"{CAFEEFAC-FFFF-FFFF-FFFF-ABCDEFFEDCBA}" = "Java Plug-in 1.6.0_31"
        #"{CAFEEFAC-FFFF-FFFF-FFFF-ABCDEFFEDCBA}" = "Java Plug-in 1.6.0_37"
        "{FCADE536-93F5-4577-80A3-E7C32FAC4C7D}" = "Loader Class v5"
        "{761497BB-D6F0-462C-B6EB-D4DAF1D92D43}" = "Java(tm) Plug-In SSV Helper"
        #"{761497BB-D6F0-462C-B6EB-D4DAF1D92D43}" = "SSVHelper Class"
        "{DBC80044-A445-435b-BC74-9C25C1C588A9}" = "Java(tm) Plug-In 2 SSV Helper"
        "{07B06095-5687-4D13-9E32-12B4259C9813}" = "STSUpld.UploadCtl"
        "{18DF081C-E8AD-4283-A596-FA578C2EBDC3}" = "Adobe PDF Link Helper"
        #"{18DF081C-E8AD-4283-A596-FA578C2EBDC3}" = "Adobe PDF Link Helper"
        "{1CC6F158-C938-424B-A757-8DC337545084}" = "Microsoft Lync Web App Plug-in"
        "{233C1507-6A77-46A4-9443-F871F945D258}" = "Shockwave ActiveX Control"
        #"{233C1507-6A77-46A4-9443-F871F945D258}" = "Shockwave ActiveX Control"
        "{3605B612-C3CF-4AB4-A426-2D853391DB2E}" = "Certificates Class"
        #"{3605B612-C3CF-4ab4-A426-2D853391DB2E}" = "Certificates Class"
        "{3640A335-73A6-424C-A6E8-B21DCCCABD0C}" = "Whale SSL Wrapper"
        "{40C37B6C-D273-41E2-8122-A338BBDB2528}" = "Microsoft Lync Web App Plug-in"
        "{424BE3CD-34AB-4F51-9C57-4341166DC8FA}" = "UCOfficeIntegration Class"
        #"{424BE3CD-34AB-4F51-9C57-4341166DC8FA}" = "(no CLSID name)"
        "{53C06A7B-FC1E-40E6-9668-31CD219BAEA7}" = "Microsoft Lync Web App Version Plug-in"
        "{611B6CB4-ACE6-4655-8D60-15FAC4AD0952}" = "Gatekeeper Class"
        "{62B4D041-4667-40B6-BB50-4BC0A5043A73}" = "SharePoint Export Database Launcher"
        "{656E5CEE-3585-4C95-AD65-037CB12288F6}" = "Forefront UAG endpoint detection"
        "{65BCBEE4-7728-41A0-97BE-14E1CAE36AAE}" = "Microsoft Office List 15.0"
        "{8AC780E1-BCDB-4816-A6EA-A88BCC064453}" = "LyncForwarder Class"
        "{9203C2CB-1DC1-482D-967E-597AFF270F0D}" = "SharePoint OpenDocuments Class"
        "{9ED13477-E909-45BC-BADC-2106D04D6BD7}" = "SharePoint DragUpload Control"
        "{A0651028-BA7A-4D71-877F-12E0175A5806}" = "UCOfficeIntegration Class"
        "{A99E6846-B0A9-4E5E-AED1-ACEA8CBEF92E}" = "Device Session Cleanup"
        "{BDEADEF5-C265-11D0-BCED-00A0C90AB50F}" = "SharePoint Stssync Handler"
        "{CA8A9780-280D-11CF-A24D-444553540000}" = "Adobe PDF Reader"
        "{E18FEC31-2EA1-49A2-A7A6-902DC0D1FF05}" = "NameCtrl Class"
        "{E7339A62-0E31-4A5E-BA3D-F2FEDFBF8BE5}" = "PersonalSite Class"
        "{00024522-0000-0000-C000-000000000046}" = "RefEdit.Ctrl"
        "{261B8CA9-3BAF-4BD0-B0C2-BF04286785C6}" = "Microsoft Outlook View Control"
        #"{261B8CA9-3BAF-4BD0-B0C2-BF04286785C6}" = "Microsoft Office Outlook View Control"
        "{3D8152C1-0CFD-4968-9684-794046886E31}" = "Microsoft Animation Control 6.0 (SP6)"
        "{9A948063-66C3-4F63-AB46-582EDAA35047}" = "Microsoft TabStrip Control 6.0 (SP6)"
        "{4D588145-A84B-4100-85D7-FD2EA1D19831}" = "Microsoft Date and Time Picker Control 6.0 (SP6)"
        "{F1651457-356D-4CA2-989D-701606A4C828}" = "Microsoft MonthView Control 6.0 (SP6)"
        "{F8CF7A98-2C45-4c8d-9151-2D716989DDAB}" = "Microsoft Visio Document"
        #"{F8CF7A98-2C45-4C8D-9151-2D716989DDAB}" = "Microsoft Visio Document"
        "{556C2772-F1AD-4DE1-8456-BD6E8F66113B}" = "Microsoft ImageList Control 6.0 (SP6)"
        "{A0E7BF67-8D30-4620-8825-7111714C7CAB}" = "Microsoft ProgressBar Control, version 6.0"
        "{CEDFFAFD-3C2F-4552-9FD3-3DC4299057FD}" = "Microsoft UpDown Control 6.0 (SP6)"
        "{585AA280-ED8B-46B2-93AE-132ECFA1DAFC}" = "Microsoft StatusBar Control 6.0 (SP6)"
        "{550C8FFB-4DC0-4756-828C-862E6D0AE74F}" = "Chain Class"
        "{8B2ADD10-33B7-4506-9569-0A1E1DBBEBAE}" = "Microsoft Toolbar Control 6.0 (SP6)"
        "{91D221C4-0CD4-461C-A728-01D509321556}" = "Store Class"
        "{95F0B3BE-E8AC-4995-9DCA-419849E06410}" = "Microsoft TreeView Control 6.0 (SP6)"
        "{CCDB0DF2-FD1A-4856-80BC-32929D8359B7}" = "Microsoft ListView Control 6.0 (SP6)"
        "{CAFEEFAC-DEC7-0000-0001-ABCDEFFEDCBA}" = "Deployment Toolkit"
        #"{CAFEEFAC-DEC7-0000-0001-ABCDEFFEDCBA}" = "Deployment Toolkit"
        "{87DACC48-F1C5-4AF3-84BA-A2A72C2AB959}" = "Microsoft ImageComboBox Control, version 6.0"
        "{9171C115-7DD9-46BA-B1E5-0ED50AFFC1B8}" = "Certificate Class"
        "{0B314611-2C19-4AB4-8513-A6EEA569D3C4}" = "Microsoft Slider Control, version 6.0"
        "{CFA7636D-CAA1-4F18-868F-8720624C8B86}" = "Microsoft Flat Scrollbar Control 6.0 (SP6)"
        "{3BCEAAF6-6774-4137-BC4E-BD8A2CD4CA95}" = "ALM Platform Loader v11.5x"
        "{6D53EC84-6AAE-4787-AEEE-F4628F01010C}" = "Symantec Intrusion Prevention"
        #"{6D53EC84-6AAE-4787-AEEE-F4628F01010C}" = "Symantec Vulnerability Protection"
        "{FF059E31-CC5A-4E2E-BF3B-96E929D65503}" = "Research"
        "{0468C085-CA5B-11D0-AF08-00609797F0E0}" = "Outlook Today's Data-binding control"
        #"{0468C085-CA5B-11D0-AF08-00609797F0E0}" = "(no CLSID name)"
        #"{0468C085-CA5B-11D0-AF08-00609797F0E0}" = "DataCtl Class"
        "{08B0E5C0-4FCB-11CF-AAA5-00401C608501}" = "(no CLSID name)"
        #"{08B0E5C0-4FCB-11CF-AAA5-00401C608501}" = "Web Browser Applet Control"
        "{88D96A0A-F192-11D4-A65F-0040963251E5}" = "XML HTTP 6.0"
        "{7466A304-ABF5-4998-88AE-F78D6F134E00}" = "ImexGridCtrl.2 Object"
        #"{7466A304-ABF5-4998-88AE-F78D6F134E00}" = "ImexGridCtrl.1 Object"
        "{444D2D27-02E8-486B-9018-3644958EF8A9}" = "FieldListCtrl.2 Object"
        #"{444D2D27-02E8-486B-9018-3644958EF8A9}" = "FieldListCtrl.1 Object"
        "{46857999-9b7c-4895-9d22-81a4a2478868}" = "Web Test Recorder 12.0"
        "{3050F819-98B5-11CF-BB82-00AA00BDCE0B}" = "HtmlDlgSafeHelper Class"
        "{432DD630-7E03-4C97-9D62-B99F52DF4FC2}" = "Microsoft Web Test Recorder 12.0 Helper"
        #"{432dd630-7e03-4c97-9d62-b99f52df4fc2}" = "Microsoft Web Test Recorder 12.0 Helper"
        "{02BCC737-B171-4746-94C9-0D8A0B2C0089}" = "Microsoft Office Template and Media Control"
        "{02BF25D5-8C17-4B23-BC80-D3488ABDDC6B}" = "QuickTime Plugin Control"
        #"{02BF25D5-8C17-4B23-BC80-D3488ABDDC6B}" = "QuickTime Object"
        "{036F8A56-0BC8-4607-8F98-D3231E6FF5ED}" = "CentraUpdaterAxCtl Class"
        "{17492023-C23A-453E-A040-C7C580BBF700}" = "Windows Genuine Advantage Validation Tool"
        "{82774781-8F4E-11D1-AB1C-0000F8773BF0}" = "DLC Class"
        "{CAFEEFAC-0016-0000-0016-ABCDEFFEDCBA}" = "Java Plug-in 1.6.0_16"
        "{E06E2E99-0AA1-11D4-ABA6-0060082AA75C}" = "GpcContainer Class"
        #"{E06E2E99-0AA1-11D4-ABA6-0060082AA75C}" = "(no CLSID name)"
        "{CF819DA3-9882-4944-ADF5-6EF17ECF3C6E}" = "Fiddler"
        "{22D6F312-B0F6-11D0-94AB-0080C74C7E95}" = "Windows Media Player"
        "{25336921-03F9-11CF-8FD0-00AA00686F13}" = "Microsoft HTML Document 6.0"
        "{2933BF94-7B36-11D2-B20E-00C04F983E60}" = "XSL Template"
        "{38481807-CA0E-42D2-BF39-B33AF135CC4D}" = "IETag Factory"
        "{39125640-8D80-11DC-A2FE-C5C455D89593}" = "Google Talk ActiveX Plugin"
        "{3FD37ABB-F90A-4DE5-AA38-179629E64C2F}" = "SharePoint Spreadsheet Launcher"
        "{4063BE15-3B08-470D-A0D5-B37161CFFD69}" = "QuickTime Plugin Control"
        #"{4063BE15-3B08-470D-A0D5-B37161CFFD69}" = "QuickTime Object"
        "{55136805-B2DE-11D1-B9F2-00A0C98BC547}" = "Shell Name Space"
        "{5852F5ED-8BF4-11D4-A245-0080C6F74284}" = "isInstalled Class"
        "{61E40D31-993D-4777-8FA0-19CA59B6D0BB}" = "Contact Selector"
        "{88D969E5-F192-11D4-A65F-0040963251E5}" = "XML DOM Document 5.0"
        "{88D969EA-F192-11D4-A65F-0040963251E5}" = "XML HTTP 5.0"
        "{88D96A05-F192-11D4-A65F-0040963251E5}" = "XML DOM Document 6.0"
        "{AB9F4455-E591-4132-A386-0B91EAEDB96C}" = "Google Talk Video Renderer"
        "{C3101A8B-0EE1-4612-BFE9-41FFC1A3C19D}" = "Google Update Plugin"
        "{C442AC41-9200-4770-8CC0-7CDB4F245C55}" = "Google Update Plugin"
        "{CD3AFA76-B84F-48F0-9393-7EDC34128127}" = "AUDIO__MP3 Moniker Class"
        "{CD3AFA99-B84F-48F0-9393-7EDC34128127}" = "VIDEO__MP4 Moniker Class"
        "{CFBFAE00-17A6-11D0-99CB-00C04FD64497}" = "Microsoft Url Search Hook"
        "{F5078F35-C551-11D3-89B9-0000F81FE221}" = "XML HTTP 3.0"
        "{F6D90F12-9C73-11D3-B32E-00C04F990BB4}" = "Free Threaded XML DOM Document"
        "{1A6FE369-F28C-4AD9-A3E6-2BCB50807CF1}" = "Developer Tools"
        #"{1A6FE369-F28C-4AD9-A3E6-2BCB50807CF1}" = "Developer Tools"
        "{72853161-30C5-4D22-B7F9-0BBC1D38A37E}" = "Groove GFS Browser Helper"
        "{7DB2D5A0-7241-4E79-B68D-6309F01C5231}" = "scriptproxy"
        "{88D96A06-F192-11D4-A65F-0040963251E5}" = "Free Threaded XML DOM Document 6.0"
        "{88D96A08-F192-11D4-A65F-0040963251E5}" = "XSL Template 6.0"
        "{AD17B774-7F87-4141-BB9C-2AEE3841DC4E}" = "Aspera Web"
        "{238F6F83-B8B4-11CF-8771-00A024541EE3}" = "Citrix ICA Client"
        "{531D5A4A-03D9-4404-AFF7-235A48E6B61E}" = "AwInstaller Class"
        "{88D969C0-F192-11D4-A65F-0040963251E5}" = "XML DOM Document 4.0"
        "{88D969C5-F192-11D4-A65F-0040963251E5}" = "XML HTTP 4.0"
        "{CB927D12-4FF7-4A9E-A169-56E4B8A75598}" = "Behavior Object"
        "{D9806E4E-82CE-4A75-83D0-A062EC605349}" = "AFContextMenuCtrl Class"
        "{DE4AF3B0-F4D4-11D3-B41A-0050DA2E6C21}" = "QuickTimeCheck Class"
        "{2BEC8FA8-1193-4A15-B8AF-C6DF6E6930C7}" = "Microsoft UpDown Control, version 5.0 (SP2)"
        "{E44F7BD4-3AB1-4D55-9190-FC53343AD2D2}" = "Microsoft TreeView Control, version 5.0 (SP2)"
        "{612685EF-57C8-469F-88AB-E4E0B595C5AB}" = "Microsoft ProgressBar Control, version 5.0 (SP2)"
        "{D8C1B55B-12DC-457F-97EC-4B84305FAA13}" = "Microsoft Hierarchical FlexGrid Control 6.0 (SP6) (OLEDB)"
        "{261399BF-4DBC-4731-B79F-EF8871D7CB36}" = "Microsoft Animation Control, version 5.0 (SP2)"
        "{1EAC2F2A-251F-4BA8-8617-99A8DD715453}" = "StdDataValue Object"
        "{2B577565-36F7-4351-B2E7-DAFC75E9D72A}" = "Microsoft Slider Control, version 5.0 (SP2)"
        "{894BA3A3-3CA3-402F-B4FE-CD08337E9535}" = "Microsoft Rich Textbox Control 6.0 (SP6)"
        "{79C784C5-8F0D-4A55-ADB3-590CCFC8EB0D}" = "Microsoft ListView Control, version 5.0 (SP2)"
        "{53749718-F78D-4A67-8703-8AE050075170}" = "Microsoft ImageList Control, version 5.0 (SP2)"
        "{97992019-74A6-46C7-9CA3-7F8C0D39940B}" = "Microsoft Toolbar Control, version 5.0 (SP2)"
        "{74DD2713-BA98-4D10-A16E-270BBEB9B555}" = "Microsoft FlexGrid Control, version 6.0 (SP6)"
        "{E8F8E80F-02EB-44CC-ABB5-6E5132BA6B24}" = "Microsoft StatusBar Control, version 5.0 (SP2)"
        "{7E96FC67-468E-4E70-B246-D42078DD2361}" = "StdDataFormat Object"
        "{D606EEC9-8368-4F10-88DB-BF5563EC36F6}" = "StdDataFormats Object"
        "{44E266A2-CD46-47A0-9ED5-EEEC5F0C2A6E}" = "Microsoft TabStrip Control, version 5.0 (SP2)"
        "{942085FD-8AEE-465F-ADD7-5E7AA28F8C14}" = "Microsoft Tabbed Dialog Control 6.0 (SP6)"
        "{225957BB-0005-48B9-8BFB-11AEE66779FB}" = "Microsoft DataGrid Control 6.0 (SP6) (OLEDB)"
        "{8F0F480A-4366-4737-8265-2AD6FDAC8C31}" = "Microsoft Common Dialog Control, version 6.0 (SP6)"
        "{8ABE89E2-1A1E-469B-8AF0-0A111727CFA5}" = "Gatekeeper Class"
        "{AA570693-00E2-4907-B6F1-60A1199B030C}" = "JuniperSetupClientControl64 Class"
        "{DF912424-425A-4F52-985D-1F83DA468AEB}" = "MeetingJoinHelper Class"
        "{773373E5-DD6A-40EB-9ED3-B16FB47F316A}" = "FileMgt.FileMgtCtrl"
        "{CAFEEFAC-0016-0000-0043-ABCDEFFEDCBA}" = "Java Plug-in 1.6.0_43"
        "{E5F5D008-DD2C-4D32-977D-1A0ADF03058B}" = "JuniperSetupControlXP Class"
        "{F27237D7-93C8-44C2-AC6E-D6057B9A918F}" = "JuniperSetupClientControl Class"
        "{64247C52-5C34-4597-B2A3-17BF5617F17F}" = "Taxonomy Control"
        "{901E885F-631B-42C8-982C-76884E5E21A0}" = "Contact Selector"
        "{CAFEEFAC-0016-0000-0024-ABCDEFFEDCBA}" = "Java Plug-in 1.6.0_24"
        "{8dcb7100-df86-4384-8842-8fa844297b3f}" = "Bing Bar"
        "{d2ce3e00-f94a-4740-988e-03dc2f38c34f}" = "Bing Bar Helper"
        "{5220CB21-C88D-11CF-B347-00AA00A28331}" = "Microsoft Licensed Class Manager 1.0"
        "{F9043C85-F6F2-101A-A3C9-08002B2F49FB}" = "Microsoft Common Dialog Control, version 6.0 (SP6)"
        "{CAFEEFAC-0015-0000-0014-ABCDEFFEDCBA}" = "Java Plug-in 1.5.0_14"
        "{CAFEEFAC-0015-0000-0014-ABCDEFFEDCBC}" = "Sun Java Console"
        "{CAFEEFAC-0015-0000-0022-ABCDEFFEDCBA}" = "Java Plug-in 1.5.0_22"
        "{CAFEEFAC-0015-0000-0022-ABCDEFFEDCBC}" = "Sun Java Console"
        "{CAFEEFAC-0017-0000-0000-ABCDEFFEDCBA}" = "Java Plug-in 1.7.0"
        "{2E5E4BAC-FEC7-4DD6-AFAF-F4139B1B9FB7}" = "LsiBrowserHook Class"
        "{95F35795-64B1-495D-9DE7-390EECC31EC0}" = "Microsoft Office Project Task Launch Control"
        "{CFC399AF-D876-11D0-9C10-00C04FC99C8E}" = "Msxml"
        "{5AE58FCF-6F6A-49B2-B064-02492C66E3F4}" = "MUCatalogWebControl Class"
        "{CAFEEFAC-0016-0000-0031-ABCDEFFEDCBA}" = "Java Plug-in 1.6.0_31"
        "{0A9CDB52-EBDF-4210-9C6A-B90C2FD410AB}" = "PowerBroker Desktops Browser Helper"
        "{2E5E4BAC-FEC7-4DD6-AFAF-F4139B1B9FB6}" = "LsiBrowserHook Class"
        "{24DA047B-40C0-4018-841B-6B7409F730FC}" = "Adobe Acrobat Sharepoint OpenDocuments Component"
        "{8075631E-5146-11D5-A672-00B0D022E945}" = "SharepointOpenXMLDocuments"
        "{36D792B3-CEA5-454E-A7EF-5B045E60EDEF}" = "DBGrid  Control"
        "{E304B70C-0FCE-4E1B-9C81-CDAAD9F7DA55}" = "Microsoft DBList Control, version 6.0"
        "{783D26D7-B4A4-4CFB-8531-78C5DCF52C8E}" = "WebClass"
        "{47DEF242-7DAF-4828-936A-895FC81D92F8}" = "Microsoft MAPI Session Control, version 6.0"
        "{1B6413C2-C55E-4BA7-B4DF-1A71DBC6ACC2}" = "Microsoft MAPI Messages Control, version 6.0 (SP6)"
        "{20E72BC7-287F-4FCD-BFB7-156FF242C27C}" = "ExportFormat Object"
        "{018BCA43-2122-4211-9589-458B6A6E2A63}" = "ExportFormats Object"
        "{6E5311A1-325D-4FFD-9AF4-B373F02AE458}" = "Microsoft WinSock Control, version 6.0 (SP6)"
        "{AFB66F3E-7A33-41E9-A4F7-FE87B64F5555}" = "Microsoft Picture Clip Control, version 6.0 (SP6)"
        "{D7FFEFBC-C693-4E6F-AE2E-ED001389CB17}" = "DataAdapter Object"
        "{62B025F5-F551-44A9-8BA8-0118EFB9127C}" = "Microsoft Chart Control 6.0 (SP6) (OLEDB)"
        "{6785E9BB-087E-4772-8CA5-3331CC3B574E}" = "Microsoft RemoteData Control, version 6.0 (SP6)"
        "{E2D211D5-11E4-4D9E-B6DB-1E902C851A49}" = "Microsoft Internet Transfer Control 6.0 (SP6)"
        "{4EE74AEC-8008-455E-AEC5-9726CF1E85BB}" = "BindingCollection Object"
        "{E9AEB8A9-DB8B-425F-8133-69CA06187353}" = "Microsoft DataRepeater Control, version 6.0"
        "{A7F31C6B-5300-47C8-A642-5AC673794C92}" = "Microsoft Data Report Runtime 6.0 (SP4)"
        "{F6565773-FA54-45E9-941C-2505E54D5710}" = "Microsoft Communications Control, version 6.0 (SP6)"
        "{234086BB-0242-46C5-B71F-5A9B961DB911}" = "Microsoft ADO Data Control 6.0 (SP6) (OLEDB)"
        "{D88A442E-9C85-48E3-A6F8-EF61C93989A0}" = "Microsoft SysInfo Control, version 6.0 (SP6)"
        "{12E15F8F-412E-4760-94E3-BE47521668BA}" = "Data Report"
        "{CB2C5FC2-C7ED-4CC1-AF07-5C5485DAB3B1}" = "DHTMLPageRuntime Object"
        "{E436987E-F427-4AD7-8738-6D0895A3E93F}" = "Addin Class"
        "{AB5357A7-3179-47F9-A705-966B8B936D5E}" = "Addin Class"
        "{1E9B270D-5829-490E-84F5-1C25D74BF01D}" = "DHTMLPageRuntimeWinEvent Object"
        "{F65348F7-505D-4FAB-B66C-D76CFFC2BD78}" = "Microsoft Multimedia Control, version 6.0 (SP6)"
        "{A57635FC-8D02-4D32-8B6E-4FBD4E2DB8A7}" = "Microsoft Masked Edit Control, version 6.0 (SP6)"
        "{D6F004C5-DC12-4B65-8730-2E95AD459F10}" = "DHTMLPageRuntimeEvent Object"
        "{E404CD92-E7B8-4037-918D-5A18CFD09ED3}" = "Microsoft DataList Control, version 6.0 (SP6) (OLEDB)"
        "{D3CCB2F7-0D00-4F26-9569-D7C368DE34E2}" = "Microsoft DataCombo Control, version 6.0 (SP6) (OLEDB)"
        "{30854451-8F2D-4282-8070-73A801B560A3}" = "Microsoft DBCombo Control, version 6.0"
        "{0D43FE01-F093-11CF-8940-00A0C9054228}" = "FileSystem Object"
        "{48123BC4-99D9-11D1-A6B3-00C04FD91555}" = "XML Document"
        "{4EB89FF4-7F78-4A0F-8B8D-2BF02E94E4B2}" = "Microsoft RDP Client Control (redistributable) - version 6"
        "{72C24DD5-D70A-438B-8A42-98424B88AFB8}" = "Windows Script Host Shell Object"
        "{F5078F40-C551-11D3-89B9-0000F81FE221}" = "XML Document 3.0"
        "{3356DB7C-58A7-11D4-AA5C-006097314BF8}" = "LaunchObj Class"
        "{38681FBD-D4CC-4A59-A527-B3136DB711D3}" = "Tumbleweed SecureTransport FileTransfer English"
        "{4871A87A-BFDD-4106-8153-FFDE2BAC2967}" = "DLM Control"
        "{99098758-CB85-4A90-924F-F21898796281}" = "Microsoft Office Slide Library Control"
        "{F9152AEC-3462-4632-8087-EEE3C3CDDA24}" = "GEPluginCoClass Object"
        "{24B224E0-9545-4A2F-ABD5-86AA8A849385}" = "Microsoft TabStrip Control, version 6.0"
        "{F91CAF91-225B-43A7-BB9E-472F991FC402}" = "Microsoft ImageList Control, version 6.0"
        "{7DC6F291-BF55-4E50-B619-EF672D9DCC58}" = "Microsoft Toolbar Control, version 6.0"
        "{627C8B79-918A-4C5C-9E19-20F66BF30B86}" = "Microsoft StatusBar Control, version 6.0"
        "{996BF5E0-8044-4650-ADEB-0B013914E99C}" = "Microsoft ListView Control, version 6.0"
        "{9181DC5F-E07D-418A-ACA6-8EEA1ECB8E9E}" = "Microsoft TreeView Control, version 6.0"
        "{12A66224-5E8A-4679-8941-0B9B960BF5EA}" = "VistaWUWebControl Class"
        "{166B1BCA-3F9C-11CF-8075-444553540000}" = "Shockwave ActiveX Control"
        "{2A646672-9C3A-4C28-9A7A-1FB0F63F28B6}" = "IE 4.x-6.x BHO for Internet Download Accelerator"
        "{9819CC0E-9669-4D01-9CD7-2C66DA43AC6C}" = "&amp;Internet Download Accelerator"
        "{0055C089-8582-441B-A0BF-17B458C2A3A8}" = "IDM integration (IDMIEHlprObj Class)"
        "{0002E541-0000-0000-C000-000000000046}" = "Microsoft Office Spreadsheet 10.0"
        "{0002E542-0000-0000-C000-000000000046}" = "Microsoft Office PivotTable 10.0"
        "{0002E543-0000-0000-C000-000000000046}" = "Microsoft Office Data Source Control 10.0"
        "{0002E546-0000-0000-C000-000000000046}" = "Microsoft Office Chart 10.0"
        "{A9667083-5060-4f44-88FB-9FF7487BBA1B}" = "Intuit QuickBooks Connector"
        "{99D1A18F-504B-4539-8AD2-9603D4F764B8}" = "HHClass Class"
        "{D20F1B09-2417-47B9-9C6A-95ABE4B98D28}" = "InstanceFinderUtil Class"
        "{6D2459CD-9AA2-48a1-A4FB-ABB8E87F4C0D}" = "AnswerWorks 5 API"
        "{2318C2B1-4965-11d4-9B18-009027A5CD4F}" = "Google Toolbar"
        "{AA58ED58-01DD-4d91-8333-CF10577473F7}" = "Google Toolbar Helper"
        "{AF69DE43-7D58-4638-B6FA-CE66B5AD205D}" = "Google Toolbar Notifier BHO"
        "{D5233FCD-D258-4903-89B8-FB1568E7413D}" = "Act.UI.InternetExplorer.Plugins.AttachFile.CAttachFile"
        "{6F431AC3-364A-478b-BBDB-89C7CE1B18F6}" = "Attach Web page to ACT! contact..."
        "{CAFEEFAC-0016-0000-0037-ABCDEFFEDCBA}" = "Java Plug-in 1.6.0_37"
        "{CD3AFA8F-B84F-48F0-9393-7EDC34128127}" = "VIDEO__X_MS_ASF Moniker Class"
        "{CD3AFA9A-B84F-48F0-9393-7EDC34128127}" = "VIDEO__QUICKTIME Moniker Class"
        }

        # Doc Mode Reasons
        $docmodereasons = @{
        "0"  = "Uninitialized"
        "1"  =  "MSHTMPAD tracetags for DRTs"
        "2"	 = 	"Session document mode supplied"
        "3"	 = 	"FEATURE_DOCUMENT_COMPATIBLE_MODE fck"
        "4"	 = 	"X-UA-Compatible meta tag"
        "5"	 = 	"X-UA-Compatible HTTP header"
        "6"	 = 	"CVList-imposed mode"
        "7"	 = 	"Native XML Parsing Mode" 
        "8"	 = 	"Toplevel QME FCK was set, and mode was determined by it"
        "9"	 = 	"Document mode is the result of the page's doctype and the browser mode"
        "10" = 	"mode supplied as a hint (not set by a rule)"
        "11" = 	"We've been constrained to a family can only have a single mode (not set by a rule)"
        "12" = 	"Webplatform version supplied; therefore align doc mode to webplatform version"  
        "13" = 	"Top level image file is set, and mode was determined by it"
        "14" = 	"Feed viewer mode determines doc mode"
        }

	# Browser State Reasons
        $browserstatereasons = @{
        "0"  = "Unitialized";
        "1"  = "Intranet sites in Compatibility View checked";
        "2"  = "Site is on Group Policy CV list";
        "3"  = "Added to the CV list by the user";
        "4"  = "X-UA-Compatible applied to page";
        "5"  = "Set by the developer toolbar";
        "6"  = "FEATURE_BROWSER_EMULATION fck";
        "7"  = "Site on MS CV list";
        "8"  = "Site on Quirks Group Policy list";
        "9"  = "MSHTMPAD override"
        "10" = "WebPlatform version supplied";
        "11" = "Site on Enterprise CV list";
        "12" = "Browser Default";
        } 


        $zones = @{
        "-1" = "INVALID"
        "0" = "LOCAL_MACHINE"
        "1" = "INTRANET"
        "2" = "TRUSTED"
        "3" = "INTERNET"
        "4" = "UNTRUSTED"
        }
    }

    Process
    {
    $compcount = $ComputerName.count
    $si=1
    $IEUrlData = @()

        ForEach ($comp in $ComputerName)
        {
            Write-Progress -Activity "Retrieving IE Telemetry URL information from $comp" -Status "Processing $si of $compcount" -PercentComplete (($si / $compcount) * 100)
            If (Test-Connection -ComputerName "$comp" -Count 1 -Quiet)
            {
                $IEUrlInfoData = Get-WMIObject -ComputerName $comp -namespace root/cimv2/IETelemetry -query "Select * from IESystemInfo where systemKey = 'SystemKey'" -ErrorAction SilentlyContinue
                If ([string]::IsNullOrEmpty($IEUrlInfoData) -eq $false)
                {
                $IEUrlInfoDataSrc = Get-WMIObject -ComputerName $comp -namespace root/cimv2/IETelemetry -Class IEURLInfo
                $IESystemInfoDataSrc = Get-WMIObject -ComputerName $comp -namespace root/cimv2/IETelemetry -Class IESystemInfo
                $IECountInfoDataSrc = Get-WMIObject -ComputerName $comp -namespace root/cimv2/IETelemetry -Class IECountInfo 
                
                    ForEach ($entry in $IEUrlInfoDataSrc)
                    {
                        $object = New-Object -TypeName PSObject
                        $object | Add-Member -MemberType NoteProperty -Name ComputerName -Value $comp
                        $object | Add-Member -MemberType NoteProperty -Name IESystemInfo -Value $IESystemInfoDataSrc
                        $object | Add-Member -MemberType NoteProperty -Name IECountInfo -Value $IECountInfoDataSrc
                        $object | Add-Member -MemberType NoteProperty -Name ActiveXGUID -Value $entry.ActiveXGUID

                            $activexdesc = @()
                            ForEach ($guid in $entry.ActiveXGUID -split ",")
                            {
                                If ([string]::IsNullOrEmpty($guid))
                                {
                                    $guid = "No ActiveX detected"
                                    $axname = "$null"
                                }
                                Else
                                {
                                    $axname = $activexlist["$guid"]
                                    If ([string]::IsNullOrEmpty($axname))
                                    {
                                        $axname = "no reference found"
                                    }
                                }
                                $object1 = New-Object -TypeName PSObject
                                $object1 | Add-Member -MemberType NoteProperty -Name URL -Value $entry.URL
                                $object1 | Add-Member -MemberType NoteProperty -Name ActiveXGUID -Value $guid
                                $object1 | Add-Member -MemberType NoteProperty -Name Description -Value $axname
                                $activexdesc += $object1
                            }
                        $object | Add-Member -MemberType NoteProperty -Name ActiveXDetail -Value $activexdesc
                        $object | Add-Member -MemberType NoteProperty -Name BrowserStateReason -Value $entry.BrowserStateReason
                        $object | Add-Member -MemberType NoteProperty -Name BrowserStateReasonDesc -Value $browserstatereasons["$($entry.BrowserStateReason)"]
                        $object | Add-Member -MemberType NoteProperty -Name CrashCount -Value $entry.CrashCount
                        $object | Add-Member -MemberType NoteProperty -Name DocMode -Value $entry.DocMode
                        $object | Add-Member -MemberType NoteProperty -Name DocModeReason -Value $entry.DocModeReason
                        $object | Add-Member -MemberType NoteProperty -Name DocModeReasonDesc -Value  $docmodereasons["$($entry.DocModeReason)"]
                        $object | Add-Member -MemberType NoteProperty -Name Domain -Value $entry.Domain
                        $object | Add-Member -MemberType NoteProperty -Name HangCount -Value $entry.HangCount
                        $object | Add-Member -MemberType NoteProperty -Name MostRecentNavigationFailure -Value $entry.MostRecentNavigationFailure
                        $object | Add-Member -MemberType NoteProperty -Name NavigationFailureCount -Value $entry.NavigationFailureCount
                        $object | Add-Member -MemberType NoteProperty -Name NumberOfVisits -Value $entry.NumberOfVisits
                        $object | Add-Member -MemberType NoteProperty -Name URL -Value $entry.URL
                        $object | Add-Member -MemberType NoteProperty -Name Zone -Value $entry.Zone
                        $object | Add-Member -MemberType NoteProperty -Name ZoneDescription -Value $zones["$($entry.Zone)"]
                        $IEUrlData += $object
                    }
                }
                Else
                {
                    Write-Output "Namespace: root/cimv2/IETelemetry or Class: IESystemInfo not found on $comp"
                }
            }
            Else
            {
                Write-verbose "Computer: $comp unreachable"   
            }
        $si++ # increase progress bar count
        }
    }

    End
    {
        if ($PSBoundParameters.ContainsKey("ActiveX"))
        {
            $IEUrlData | Select-Object * | Select-Object -ExpandProperty ActiveXDetail
        }
        Else
        {
            $IEUrlData
        }
    }
}
