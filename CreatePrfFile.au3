 #include <Date.au3>

Func CreatePrfFile ($sPrfFilePath, $sPstPath, $sPopServer, $sSmtpServer, $sEmail)
    Local $sNow = _NowDate()
    Local $iEpoch = _DateDiff('s', "1970/01/01 00:00:00", _NowCalc())
    $sPrfContent = "[General]" & @CRLF & _
        ";Silent=Yes" & @CRLF & _
        "Custom=1" & @CRLF & _
        "ProfileName=PSTExporter_" & $sEmail & @CRLF & _
        "DefaultProfile=Yes" & @CRLF & _
        "OverwriteProfile=Append" & @CRLF & _
        "BackupProfile=No" & @CRLF & _
        "ModifyDefaultProfileIfPresent=false" & @CRLF & _
        "DefaultStore=Service1" & @CRLF & @CRLF & _
        "[Service List]" & @CRLF & _
        ";ServiceX=Microsoft Outlook Client" & @CRLF & _
        "ServiceEGS=Exchange Global Section" & @CRLF & _
        "Service1=Unicode Personal Folders" & @CRLF & @CRLF & _
        "[Internet Account List]" & @CRLF & _
        "Account1=I_Mail" & @CRLF & @CRLF & _
        "[ServiceEGS]" & @CRLF & @CRLF & _
        "[Service1]" & @CRLF & _
        "UniqueService=No" & @CRLF & _
        "Name=" & $sEmail & @CRLF & _
        "PathAndFilenameToPersonalFolders=" & $sPstPath & "\" & $sEmail "_" & $sNow & "_" & $iEpoch & ".pst" & @CRLF & _
        "EncryptionType=0x50000000" & @CRLF & @CRLF & _
        "[Account1]" & @CRLF & _
        "UniqueService=No" & @CRLF & _
        "AccountName=" & $sEmail & @CRLF & _
        "POP3Server=" & $sPopServer & @CRLF & _
        "SMTPServer=" & $sSmtpServer & @CRLF & _
        "POP3UserName=" & $sEmail & @CRLF & _
        "EmailAddress=" & $sEmail & @CRLF & _
        "POP3UseSPA=0" & @CRLF & _
        "DisplayName=" & @CRLF & _
        "ReplyEMailAddress=" & @CRLF & _
        "SMTPUseAuth=0" & @CRLF & _
        "SMTPAuthMethod=0" & @CRLF & _
        "ConnectionType=0" & @CRLF & _
        "LeaveOnServer=0x50001" & @CRLF & _
        "POP3UseSSL=0" & @CRLF & _
        "ConnectionOID=MyConnection" & @CRLF & _
        "POP3Port=110" & @CRLF & _
        "ServerTimeOut=60" & @CRLF & _
        "SMTPPort=25" & @CRLF & _
        "SMTPSecureConnection=0" & @CRLF & _
        "DefaultAccount=TRUE" & @CRLF & @CRLF & _
        "[Microsoft Exchange Server]" & @CRLF & _
        "ServiceName=MSEMS" & @CRLF & _
        "MDBGUID=5494A1C0297F101BA58708002B2A2517" & @CRLF & _
        "MailboxName=PT_STRING8,0x6607" & @CRLF & _
        "HomeServer=PT_STRING8,0x6608" & @CRLF & _
        "OfflineAddressBookPath=PT_STRING8,0x660E" & @CRLF & _
        "OfflineFolderPathAndFilename=PT_STRING8,0x6610" & @CRLF & @CRLF & _
        "[Exchange Global Section]" & @CRLF & _
        "SectionGUID=13dbb0c8aa05101a9bb000aa002fc45a" & @CRLF & _
        "MailboxName=PT_STRING8,0x6607" & @CRLF & _
        "HomeServer=PT_STRING8,0x6608" & @CRLF & _
        "RPCoverHTTPflags=PT_LONG,0x6623" & @CRLF & _
        "RPCProxyServer=PT_UNICODE,0x6622" & @CRLF & _
        "RPCProxyPrincipalName=PT_UNICODE,0x6625" & @CRLF & _
        "RPCProxyAuthScheme=PT_LONG,0x6627" & @CRLF & _
        "CachedExchangeConfigFlags=PT_LONG,0x6629" & @CRLF & @CRLF & _
        "[Microsoft Mail]" & @CRLF & _
        "ServiceName=MSFS" & @CRLF & _
        "ServerPath=PT_STRING8,0x6600" & @CRLF & _
        "Mailbox=PT_STRING8,0x6601" & @CRLF & _
        "Password=PT_STRING8,0x67f0" & @CRLF & _
        "RememberPassword=PT_BOOLEAN,0x6606" & @CRLF & _
        "ConnectionType=PT_LONG,0x6603" & @CRLF & _
        "UseSessionLog=PT_BOOLEAN,0x6604" & @CRLF & _
        "SessionLogPath=PT_STRING8,0x6605" & @CRLF & _
        "EnableUpload=PT_BOOLEAN,0x6620" & @CRLF & _
        "EnableDownload=PT_BOOLEAN,0x6621" & @CRLF & _
        "UploadMask=PT_LONG,0x6622" & @CRLF & _
        "NetBiosNotification=PT_BOOLEAN,0x6623" & @CRLF & _
        "NewMailPollInterval=PT_STRING8,0x6624" & @CRLF & _
        "DisplayGalOnly=PT_BOOLEAN,0x6625" & @CRLF & _
        "UseHeadersOnLAN=PT_BOOLEAN,0x6630" & @CRLF & _
        "UseLocalAdressBookOnLAN=PT_BOOLEAN,0x6631" & @CRLF & _
        "UseExternalToHelpDeliverOnLAN=PT_BOOLEAN,0x6632" & @CRLF & _
        "UseHeadersOnRAS=PT_BOOLEAN,0x6640" & @CRLF & _
        "UseLocalAdressBookOnRAS=PT_BOOLEAN,0x6641" & @CRLF & _
        "UseExternalToHelpDeliverOnRAS=PT_BOOLEAN,0x6639" & @CRLF & _
        "ConnectOnStartup=PT_BOOLEAN,0x6642" & @CRLF & _
        "DisconnectAfterRetrieveHeaders=PT_BOOLEAN,0x6643" & @CRLF & _
        "DisconnectAfterRetrieveMail=PT_BOOLEAN,0x6644" & @CRLF & _
        "DisconnectOnExit=PT_BOOLEAN,0x6645" & @CRLF & _
        "DefaultDialupConnectionName=PT_STRING8,0x6646" & @CRLF & _
        "DialupRetryCount=PT_STRING8,0x6648" & @CRLF & _
        "DialupRetryDelay=PT_STRING8,0x6649" & @CRLF & @CRLF & _
        "[Personal Folders]" & @CRLF & _
        "ServiceName=MSPST MS" & @CRLF & _
        "Name=PT_STRING8,0x3001" & @CRLF & _
        "PathAndFilenameToPersonalFolders=PT_STRING8,0x6700 " & @CRLF & _
        "RememberPassword=PT_BOOLEAN,0x6701" & @CRLF & _
        "EncryptionType=PT_LONG,0x6702" & @CRLF & _
        "Password=PT_STRING8,0x6703" & @CRLF & @CRLF & _
        "[Unicode Personal Folders]" & @CRLF & _
        "ServiceName=MSUPST MS" & @CRLF & _
        "Name=PT_UNICODE,0x3001" & @CRLF & _
        "PathAndFilenameToPersonalFolders=PT_STRING8,0x6700 " & @CRLF & _
        "RememberPassword=PT_BOOLEAN,0x6701" & @CRLF & _
        "EncryptionType=PT_LONG,0x6702" & @CRLF & _
        "Password=PT_STRING8,0x6703" & @CRLF & @CRLF & _
        "[Outlook Address Book]" & @CRLF & _
        "ServiceName=CONTAB" & @CRLF & @CRLF & _
        "[LDAP Directory]" & @CRLF & _
        "ServiceName=EMABLT" & @CRLF & _
        "ServerName=PT_STRING8,0x6600" & @CRLF & _
        "UserName=PT_STRING8,0x6602" & @CRLF & _
        "UseSSL=PT_BOOLEAN,0x6613" & @CRLF & _
        "UseSPA=PT_BOOLEAN,0x6615" & @CRLF & _
        "EnableBrowsing=PT_BOOLEAN,0x6622" & @CRLF & _
        "DisplayName=PT_STRING8,0x3001" & @CRLF & _
        "ConnectionPort=PT_STRING8,0x6601" & @CRLF & _
        "SearchTimeout=PT_STRING8,0x6607" & @CRLF & _
        "MaxEntriesReturned=PT_STRING8,0x6608" & @CRLF & _
        "SearchBase=PT_STRING8,0x6603" & @CRLF & _
        "CheckNames=PT_STRING8,0x6624" & @CRLF & _
        "DefaultSearch=PT_LONG,0x6623" & @CRLF & @CRLF & _
        "[Microsoft Outlook Client]" & @CRLF & _
        "SectionGUID=0a0d020000000000c000000000000046" & @CRLF & _
        "FormDirectoryPage=PT_STRING8,0x0270" & @CRLF & _
        "WebServicesLocation=PT_STRING8,0x0271" & @CRLF & _
        "ComposeWithWebServices=PT_BOOLEAN,0x0272" & @CRLF & _
        "PromptWhenUsingWebServices=PT_BOOLEAN,0x0273" & @CRLF & _
        "OpenWithWebServices=PT_BOOLEAN,0x0274" & @CRLF & _
        "CachedExchangeMode=PT_LONG,0x041f" & @CRLF & _
        "CachedExchangeSlowDetect=PT_BOOLEAN,0x0420" & @CRLF & @CRLF & _
        "[Personal Address Book]" & @CRLF & _
        "ServiceName=MSPST AB" & @CRLF & _
        "NameOfPAB=PT_STRING8,0x001e3001" & @CRLF & _
        "PathAndFilename=PT_STRING8,0x001e6600" & @CRLF & _
        "ShowNamesBy=PT_LONG,0x00036601" & @CRLF & @CRLF & _
        "[I_Mail]" & @CRLF & _
        "AccountType=POP3" & @CRLF & _
        ";--- POP3 Account Settings ---" & @CRLF & _
        "AccountName=PT_UNICODE,0x0002" & @CRLF & _
        "DisplayName=PT_UNICODE,0x000B" & @CRLF & _
        "EmailAddress=PT_UNICODE,0x000C" & @CRLF & _
        ";--- POP3 Account Settings ---" & @CRLF & _
        "POP3Server=PT_UNICODE,0x0100" & @CRLF & _
        "POP3UserName=PT_UNICODE,0x0101" & @CRLF & _
        "POP3UseSPA=PT_LONG,0x0108" & @CRLF & _
        "Organization=PT_UNICODE,0x0107" & @CRLF & _
        "ReplyEmailAddress=PT_UNICODE,0x0103" & @CRLF & _
        "POP3Port=PT_LONG,0x0104" & @CRLF & _
        "POP3UseSSL=PT_LONG,0x0105" & @CRLF & _
        "; --- SMTP Account Settings ---" & @CRLF & _
        "SMTPServer=PT_UNICODE,0x0200" & @CRLF & _
        "SMTPUseAuth=PT_LONG,0x0203" & @CRLF & _
        "SMTPAuthMethod=PT_LONG,0x0208" & @CRLF & _
        "SMTPUserName=PT_UNICODE,0x0204" & @CRLF & _
        "SMTPUseSPA=PT_LONG,0x0207" & @CRLF & _
        "ConnectionType=PT_LONG,0x000F" & @CRLF & _
        "ConnectionOID=PT_UNICODE,0x0010" & @CRLF & _
        "SMTPPort=PT_LONG,0x0201" & @CRLF & _
        "SMTPSecureConnection=PT_LONG,0x020A" & @CRLF & _
        "ServerTimeOut=PT_LONG,0x0209" & @CRLF & _
        "LeaveOnServer=PT_LONG,0x1000" & @CRLF & @CRLF & _
        "[IMAP_I_Mail]" & @CRLF & _
        "AccountType=IMAP" & @CRLF & _
        ";--- IMAP Account Settings ---" & @CRLF & _
        "AccountName=PT_UNICODE,0x0002" & @CRLF & _
        "DisplayName=PT_UNICODE,0x000B" & @CRLF & _
        "EmailAddress=PT_UNICODE,0x000C" & @CRLF & _
        ";--- IMAP Account Settings ---" & @CRLF & _
        "IMAPServer=PT_UNICODE,0x0100" & @CRLF & _
        "IMAPUserName=PT_UNICODE,0x0101" & @CRLF & _
        "IMAPUseSPA=PT_LONG,0x0108" & @CRLF & _
        "Organization=PT_UNICODE,0x0107" & @CRLF & _
        "ReplyEmailAddress=PT_UNICODE,0x0103" & @CRLF & _
        "IMAPPort=PT_LONG,0x0104" & @CRLF & _
        "IMAPUseSSL=PT_LONG,0x0105" & @CRLF & _
        "; --- SMTP Account Settings ---" & @CRLF & _
        "SMTPServer=PT_UNICODE,0x0200" & @CRLF & _
        "SMTPUseAuth=PT_LONG,0x0203" & @CRLF & _
        "SMTPAuthMethod=PT_LONG,0x0208" & @CRLF & _
        "SMTPUserName=PT_UNICODE,0x0204" & @CRLF & _
        "SMTPUseSPA=PT_LONG,0x0207" & @CRLF & _
        "ConnectionType=PT_LONG,0x000F" & @CRLF & _
        "ConnectionOID=PT_UNICODE,0x0010" & @CRLF & _
        "SMTPPort=PT_LONG,0x0201" & @CRLF & _
        "SMTPSecureConnection=PT_LONG,0x020A" & @CRLF & _
        "ServerTimeOut=PT_LONG,0x0209" & @CRLF & _
        "CheckNewImap=PT_LONG,0x1100" & @CRLF & _
        "RootFolder=PT_UNICODE,0x1101" & @CRLF & @CRLF & _
        "[INET_HTTP]" & @CRLF & _
        "AccountType=HOTMAIL" & @CRLF & _
        "Account=PT_UNICODE,0x0002" & @CRLF & _
        "HttpServer=PT_UNICODE,0x0100" & @CRLF & _
        "UserName=PT_UNICODE,0x0101" & @CRLF & _
        "Organization=PT_UNICODE,0x0107" & @CRLF & _
        "UseSPA=PT_LONG,0x0108" & @CRLF & _
        "TimeOut=PT_LONG,0x0209" & @CRLF & _
        "Reply=PT_UNICODE,0x0103" & @CRLF & _
        "EmailAddress=PT_UNICODE,0x000C" & @CRLF & _
        "FullName=PT_UNICODE,0x000B" & @CRLF & _
        "Connection Type=PT_LONG,0x000F" & @CRLF & _
        "ConnectOID=PT_UNICODE,0x0010" & @CRLF & _

    return FileWrite($sPrfFilePath, $sPrfContent)
EndFunc