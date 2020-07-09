################################################################ DEBUT DES GLOBALES ####################################################################################
#$SHARECOPTER_FOLDER = "\\garlaban\fdat1010\Database Production\Generated data\Input_files\"
[String] $SHARECOPTER_FOLDER = "C:\Users\the_w\Desktop\airbus\"
[String] $SHARECOPTER_LIST = $SHARECOPTER_FOLDER + "sharecopterList.xlsx" 

#$CSV_EXTRACT = "\\garlaban\fdat1010\Database Production\Generated data\Input_files\extract.csv"
[String] $CSV_EXTRACT = "C:\Users\the_w\Desktop\airbus\extract.csv"
[String] $CSV_DELIMITER = ";"

[String] $VAR_CYCLE_AIRAC = 2005

[String] $VAR_BODY = "
        <HTML><HEAD>
            <BODY style='font-family:Calibri;'>
            
                <p><FONT size = 4>Dear valued customer,</p>
                <p>
                    <span>Your Aeronautical Data Service Databases for <b>AIRAC cycle  ###VAR_CYCLE_AIRAC### </b> are now available on ADS website. Attached you can also find the <b> Airbus certificate of conformity.</b>
                </p>

                <p>
                    <span>To access your databases, you can now logon to your <b>Airbus World account</b>, go to <b>"“Flight Ops”"</b> tab and click on <b>"“Aeronautical Data Service”".</b></span>
                    <span>If you can’t connect via Airbus World, <b>please logon via</b> <a href =""http://sharecopter-ecm-basic.eurocopter.com/sites/spsa1893/Customer/SitePages/ads.aspx""> https://sharecopter-ecm-basic.eurocopter.com/sites/spsa1893/customer </a></span>
                    <span>with your "“first”" Airbus World login which start with Z0000 xx).</span>
                </p>
 
                <p>
                    <span>Please visit your website to find complementary information about databases.</span>
                    <span>Learn how to retrieve and load Aeronautical Data for Helionix with our video tutorial available on Youtube: <a href =""http://youtu.be/fh_tU6sTH38""> http://youtu.be/fh_tU6sTH38 </a></span>
                </p>

                <p>
		            <span>If you could please fill out the customer satisfaction form available to you in the News section of the Customer Shareopter Homepage.</span>
		            <span>Thanks in advance.</span>
		        </p>

                <p>For questions or concerns, please contact your Account Executive.</p>
                <p>We sincerely hope that you and your relatives are safe during this crisis.</p>
                <p>Thank you.</p>
                <p>Best regards</p>

                <p><span style='color:#1F497D'><i><b>A</b>eronautical <b>D</b>ata <b>S</b>ervices</i><br/></span></p>
                <img src =""###CSV_EXTRACT###""></FONT>
        
            </BODY>
        </HEAD></HTML>
"

[String] $VAR_CC = "support.ads.ah@airbus.com"

$global:MAIL = $null 

################################################################ FIN DES GLOBALES ######################################################################################



################################################################  BEGIN FUNCTION UTILITAIRE #############################################################################

function CheckCar([Char] $caractere) {
    if ($caractere -match "[a-zA-Z0-9!#$%&'*+-/=^_`.{|}~@]"){
        Return $true
    }Else{
        Return $false
    }
}

function EncodeEmail([String] $mail) {
    $login=[system.text.Encoding]::UTF8.GetString([System.Text.ASCIIEncoding]::Convert( [system.text.encoding]::default  , [system.text.encoding]::UTF8, [system.text.Encoding]::Default.GetBytes($mail)))
    [int] $charIndex=0

    $charIndex=$login.Length - 1
    
    [bool]$modif=$false

    While ($charIndex -ge 0){
        If (-Not (CheckCar($login[$charIndex]))){
            $login=$login.Remove($charIndex, 1)
            $modif=$true
        }

        $charIndex--
    }

    If ($modif){
        Write-Host "Modification adresse email"
        Write-Host $mail
        Write-Host $login
    }

    return ($login)
}

function replace_var($str, $old, $new) {
    $r = $str -replace $old,$new
    return $r
}


################################################################  END FUNCTION UTILITAIRE ##########################################################################

################################################################ BEGIN FUNCTION METIER #############################################################################

function Send-Email() {
    #$global:MAIL.Display : ouvre le mail mais ne l'envoie pas ! 

    Write-Host "Send E-mail"
    $global:MAIL.Display()
}

function extract_sharecopter() {
    #Pour récupérer les données : convertir le .xlsx en .csv
    $ExcelWB = new-object -ComObject excel.application

    if(![System.IO.File]::Exists($SHARECOPTER_LIST)){
        Write-Host "Le pdf $SHARECOPTER_LIST n'existe pas"
        return
    }

    $Workbook = $ExcelWB.Workbooks.Open($SHARECOPTER_LIST)
    $Workbook.SaveAs($CSV_EXTRACT,6)
    $ExcelWB.Quit()
    Stop-Process -Name EXCEL

    if(![System.IO.File]::Exists($CSV_EXTRACT)){
        Write-Host "Le pdf $CSV_EXTRACT n'existe pas"
        return
    }

    return $(Import-Csv -Path $CSV_EXTRACT -Delimiter $CSV_DELIMITER -Encoding OEM)
}

function add_pj ([String] $pj) {
    If($pj -eq $null){ return }

    if(![System.IO.File]::Exists($pj)){
        Write-Host "Le pdf $pj n'existe pas"
        return 
    }

    $extn = [IO.Path]::GetExtension($pj)
    if(! $extn -eq ".pdf"){
        Write-Host "Le pdf $pj n'est pas un PDF"
        return 
    }

    Write-Host "Ajout de la piece jointe : $pj au mail"
    $global:MAIL.Attachments.Add($pj)  
}

function concat_destinataire ([String] $current, [String] $destinataire) {
    If($current -eq $null -and $destinataire -eq $null){ return $null }

    If($current -eq $null -and !$destinataire -eq $null){ return $destinataire  + ";"}
    If(!$current -eq $null -and $destinataire -eq $null){ return $current  + ";" }

    return $current + $destinataire + ";"
}

function add_destinataire ([String] $destinataire) {
    Write-Host "Liste des destinataire du mail : $destinataire"

    $global:MAIL.Recipients.Add($destinataire)
}

function add_cc ([String] $cc) {
    If($cc -eq $null){ return }

    $global:MAIL.CC = $cc
}

function add_body ([String] $body) {
    $body = replace_var -str $body -old "###VAR_CYCLE_AIRAC###" -new $VAR_CYCLE_AIRAC
    $body = replace_var -str $body -old "###CSV_EXTRACT###"     -new $CSV_EXTRACT

    $global:MAIL.HTMLBody = $body
}

function add_subject () {
    $global:MAIL.Subject = "AIRAC Cycle $VAR_CYCLE_AIRAC Airbus Helicopters Database available"
}

function create_empty_mail () {
    #Prépare l'objet Outlook
    $outlook = New-Object -comObject Outlook.Application
    $global:MAIL = $outlook.CreateItem(0) 
}

function create_mail([String] $destinataire, [String] $company) {
    create_empty_mail
    add_body -body $VAR_BODY
    add_subject
    add_cc -cc $VAR_CC
    add_destinataire -destinataire $destinataire

    $folderCoC=$SHARECOPTER_FOLDER + "$company"
    $contenersCOC = Get-ChildItem -recurse $folderCoC
    foreach($contenerCOC in $contenersCOC){
        add_pj -pj $contenerCOC.FullName  
    }
}

function move_coc($file, [String] $destination) {
    Write-Host "On deplace le fichier $file vers $destination"
    Move-Item -Path $file -Destination $destination
}

function create_folder_company([String] $folder) {
    #créé un dossier avec le nom du client 
    
    if([System.IO.Directory]::Exists($folder)){
        #Write-Host "Le dossier $folder existe déjà"
        return $folder
    }

    #Write-Host "Creation du dossier : $folder"
    New-Item -ItemType Directory -Force -Path $folder
}

function traitement_coc () {
    #Pour chaque PDF
    $conteners = $SHARECOPTER_FOLDER

    # $conteners = $CoCDepot
    $contenersCOC = Get-ChildItem -recurse $conteners

    #Si le CoC est un pdf
    foreach($contenerCOC in $contenersCOC){

        $CurrentCoC = $contenerCOC.Name
        if($CurrentCoC -match ".pdf"){
            [String] $base = $contenerCOC.BaseName
            [String] $companyname  = $base.Split("_")[2]
            [String] $folder = $SHARECOPTER_FOLDER + $companyname 

            create_folder_company -folder $folder
            move_coc -file $contenerCOC.FullName -destination $folder 
        }
    }
}

function traitement_mail() {
    $Sharecopters = extract_sharecopter

    foreach ($sharecopter in $Sharecopters){ 

        [String] $Company_name = $sharecopter."Company Name"
        [String] $Status       = $sharecopter."Status"

        [String] $destinataire = concat_destinataire -current $null -destinataire $(EncodeEmail($sharecopter."email address #1"))
        [String] $destinataire = concat_destinataire -current $destinataire -destinataire $(EncodeEmail($sharecopter."email address #2"))
        [String] $destinataire = concat_destinataire -current $destinataire -destinataire $(EncodeEmail($sharecopter."email address #3"))
        [String] $destinataire = concat_destinataire -current $destinataire -destinataire $(EncodeEmail($sharecopter."email address #4"))
        [String] $destinataire = concat_destinataire -current $destinataire -destinataire $(EncodeEmail($sharecopter."email address #5"))
        [String] $destinataire = concat_destinataire -current $destinataire -destinataire $(EncodeEmail($sharecopter."email address #6"))

        create_mail -destinataire $destinataire -company $Company_name

        if ($Status -match "Resolved") {
            
            #Pour chaque ligne de l'extract sharecopter
            $Folders = Get-ChildItem -Path $SHARECOPTER_FOLDER -Directory
            foreach ($MyFolder in $Folders){
                $folder=$MyFolder.name
                    
                if ($Company_name -Like $folder){
                  
                    write-Host "Starting Send-MailViaOutlook Script." 

                    #On prépare le mail                         
                    Send-Email           
                    $global:MAIL = $null;

                    #$contenersfoldercurrent = $SHARECOPTER_FOLDER + $folder
                    #Remove-Item –path $contenersfoldercurrent -recurse

                    Write-Host "Closing Send-MailViaOutlook Script." 
                }
            }
        }
    }

    remove-item $CSV_EXTRACT
}

################################################################ FIN FUNCTION METIER ##################################################################################

################################################################ MAIN FUNCTION #########################################################################################

function main {
    If (Get-Process | Where-Object { $_.ProcessName -eq "OUTLOOK"}){

        [string] $stop_outlook = $(read-host "Outlook est en cours d'exécution. Arrêt du processus Outlook ? oui/non")

        If($stop_outlook -eq "oui"){
            Stop-Process -Name OUTLOOK
        }
    }
    
    traitement_coc
    traitement_mail
}
################################################################ END MAIN FUNCTION ######################################################################################

main
