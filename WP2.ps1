######################################################################
# Frédéric VERNET
######################################################################

################################################################ DEBUT DES GLOBALES ####################################################################################

[String] $global:TemplateCoC =  "C:\Users\the_w\Desktop\airbus\Wp2 création de coc\Indus\Templates\coc_template.docx"

[String] $InputCoC   = "C:\Users\the_w\Desktop\airbus\Wp2 création de coc\Input_files\"
[String] $ExportCSV  = "C:\Users\the_w\Desktop\airbus\Wp2 création de coc\Input_files\"
[String] $ExportXlsx = "C:\Users\the_w\Desktop\airbus\Wp2 création de coc\Input_files\Export.xlsx"

[String] $WorkingDirectory = "C:\Users\the_w\Desktop\airbus\Wp2 création de coc\Input_files"
[String] $terrain = "TERRAIN"
[String] $WP2_GARLABAN_ROOTDIR = "C:\Users\the_w\Desktop\airbus\Wp2 création de coc\"
[String] $FicLogWP2="C:\Users\the_w\Desktop\airbus\Wp2 création de coc\Input_files\WP2_log.csv"

[int] $myCycleAIRAC = 2008

$global:coc_values = $null
$global:sharecopter_columns = $null

param([Parameter(Mandatory=$true, ValueFromPipeline=$false)]
      [String] $myCycleAIRAC)

Set-PSDebug -Strict

######################################################################
# Récupération des variables d'environnement
######################################################################

If (Test-Path -Path $FicLogWP2)
{
    Remove-Item $FicLogWP2
}

######################################################################
# Fonction de remplissage du document word
######################################################################
function readWord([string] $name, [string] $SignatureFile){

    $Word = New-Object -com word.application
    $Word.visible = $False

    $OpenDoc = $Word.documents.Open($global:TemplateCoC)
    $Selection = $Word.selection

    $Signature="$SignaturePath\$SignatureFile"
    # Récupération du tableau du template
    $Selection.Tables | ForEach-Object {   
        $objTables = $_      
    }

    # Insertion de la bonne signature au bon endroit
    # $objTables.Cell(20,1).Range.InlineShapes.AddPicture($Signature)
    
    $global:coc_values.GetEnumerator() | % {
        #Write-Host $_.key
        #Write-Host $_.Value
        $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
    }
      
    $chemin = $InputCoC + $name 

    write-host $chemin
    
    # Sauvegarde du CoC en docx avec le nom adéquat
    $OpenDoc.SaveAs([ref]$chemin)
    $OpenDoc.close();
    $Word = $null

    Stop-Process -name WINWORD -Force
}

######################################################################
# Fonction qui permet de convertir les CoC qui sont en .docx en .pdf
# puis supprime tous les CoC en .docx
######################################################################
function ConvertDocxtoPDF{
    $CoC = Get-ChildItem -recurse $global:InputCoC

    foreach ($CoCs in $CoC){
    # Pour convertir seulement les docx
        if($CoCs -match "docx") {
            Write-host("Conversion du CoC ", $CoCs)
            #Conversion des CoC en .pdf
            $Word = new-object -ComObject Word.Application
            $Word.Visible = $false
            $Doc=$Word.Documents.Open($CoCs.FullName)
            $Doc.saveas([ref] (($CoCs.FullName).replace(“docx”,”pdf”)), [ref] 17)
            $Doc.close()
            $Word.Quit()
            $Word = $null
            Stop-Process -name WINWORD
        }
    }
}

######################################################################
# Fonction qui permet de choisir la signature à appliquer au CoC
# Retour : la signature selectionné
######################################################################
function ChoixSignature{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

    return $null

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Select à Signature"
    $objForm.Size = New-Object System.Drawing.Size(300,200) 
    $objForm.StartPosition = "CenterScreen"

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
        {$x=$objListBox.SelectedItem;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
        {$objForm.Close()}})

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(75,120)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$global:SignatureFile=$objListBox.SelectedItem;$objForm.Close()})

    $objForm.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(150,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({$objForm.Close()})
    $objForm.Controls.Add($CancelButton)

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,20) 
    $objLabel.Size = New-Object System.Drawing.Size(280,20) 
    $objLabel.Text = "Please select a Signature:"
    $objForm.Controls.Add($objLabel) 

    $objListBox = New-Object System.Windows.Forms.ListBox 
    $objListBox.Location = New-Object System.Drawing.Size(10,40) 
    $objListBox.Size = New-Object System.Drawing.Size(260,20) 
    $objListBox.Height = 80

    foreach($file in Get-ChildItem -Path $SignaturePath\*.bmp -Recurse)
    {
        [void] $objListBox.Items.Add($file.Name)
    }

    $objForm.Controls.Add($objListBox) 

    $objForm.Topmost = $True

    $objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()

    return $SignatureFile
}

######################################################################
# Fonction de recherche dans le fichier Export.xlsx le nom du fichier 
# dans garlaban correspondant à la référence passée en paramètre 
# Entrée : $value --> référence du champs 1 du coc. Ex : 4DE1H01
# Retour : le nom du fichier correspondant
######################################################################
function ResearchFile{
    param ([string] $value)

    [string] $res = ""
    $ExcelWB = new-object -comobject excel.application 
    #Ouverture du fichier Excel Export.xlsx et convertion en fichier excel Export.csv

    $Workbook = $ExcelWB.Workbooks.Open($ExportXlsx) 
    $save = $ExportCSV + "Export.csv"
    $Workbook.SaveAs($save,6)
  
    $file = "File"
    $Export = Import-Csv -Path $save -Delimiter ";"

    foreach ($Exports in $Export){
        if ($value -match $Exports.$file.substring(0,4) ){
            $res = $Exports.$file
        } 
    }

    $Workbook.Close($false)
    $ExcelWB.quit()

    Remove-Item $save
    Stop-Process -name EXCEL
    return $res
}

######################################################################
# Fonction de conversion d'un string en date
# Entrée : $date --> date au format jj-mois-aa
# Retour : une date au format JJ/MM/YY
######################################################################
function ConvertStringToDate 
{
    param([string]$date)

    [string] $res = ""
    
    if($date.Substring(0,6) -match "jan") 
    {
        $res = "01/" + $date.Substring(0,2) + "/" + $date.Substring(7,2)
    } 

    if($date.Substring(0,6) -match "feb") 
    {
        $res = "02/" + $date.Substring(0,2) + "/" + $date.Substring(7,2)
    } 

    if($date.Substring(0,6) -match "mar") 
    {
        $res = "03/" + $date.Substring(0,2) + "/" + $date.Substring(7,2)
    } 
    if(($date.Substring(0,6) -match "avr") -or ($date.Substring(0,6) -match "apr"))
    {
        $res = "04/" + $date.Substring(0,2) + "/" + $date.Substring(7,2)
    } 

    if(($date.Substring(0,6) -match "may") -or ($date.Substring(0,6) -match "mai"))
    {
        $res = "05/" + $date.Substring(0,2) + "/" + $date.Substring(7,2)
    } 

    if($date.Substring(0,7) -match "juin") 
    {
        $res = "06/" + $date.Substring(0,2) + "/" + $date.Substring(8,2)
    } 
    if ($date.Substring(0,7) -match "jun")
    {
        $res = "06/" + $date.Substring(0,2) + "/" + $date.Substring(7,2)
    }

    if($date.Substring(0,7) -match "juil")
    {
        $res =  "07/" + $date.Substring(0,2) + "/" +  $date.Substring(8,2)
    } 
    if($date.Substring(0,7) -match "jul")
    {
        $res =  "07/" + $date.Substring(0,2) + "/" +  $date.Substring(7,2)
    } 

    if($date.Substring(0,6) -match "aou")
    {
        $res = "08/" + $date.Substring(0,2) + "/" + $date.Substring(7,2)
    } 

    if($date.Substring(0,6) -match "aug")
    {
        $res = "08/" + $date.Substring(0,2) + "/" + $date.Substring(7,2)
    } 

    if($date.Substring(0,6) -match "sep") 
    {
        $res = "09/" + $date.Substring(0,2) + "/" + $date.Substring(7,2)
    } 

    if($date.Substring(0,6) -match "oct") 
    {
        $res = "10/" + $date.Substring(0,2) + "/" + $date.Substring(7,2)
    } 

    if($date.Substring(0,6) -match "nov") 
    {
        $res = "11/" + $date.Substring(0,2) + "/" + $date.Substring(7,2)
    } 

    if($date.Substring(0,6) -match "dec") 
    {
        $res = "12/" + $date.Substring(0,2) + "/" + $date.Substring(7,2)
    } 

    $test=$date.Substring(0,6)
    
    return ($res)
}

######################################################################
# Fonction de chargement du contenu d'un fichier Excel en mémoire
# Entrée : $file_name --> nom du fichier Excel à charger
#          $colonnes_extraire --> liste des colonnes à extraire dans le fichier Excel
#          $ligne_depart --> nombre de lignes à ignorer en début de fichier
#          $ligne_fin --> nombre de lignes à ignoer en fin de fichier
# Retour : un tableau contenant l'extraction des données du fichier Excel
######################################################################
function Load_Sharecopter_In_Memory([string] $file_name){

    [int] $ligne_depart = 2

    $Excel = New-Object -ComObject excel.application 
    $Excel.visible = $False

    # Ouverture du fichier Excel et accès à la feuille numéro 1
    $Workbook = $excel.Workbooks.open($file_name) 
    $Worksheet = $Workbook.WorkSheets.Item(1)
    $Worksheet.activate()

    # Calcul du nombre de lignes
    [int] $nombre_lignes = $Worksheet.UsedRange.Rows.Count
    [int] $nombre_colonnes = $Worksheet.UsedRange.Columns.Count

    write-host ("Total Lignes : ",$nombre_lignes)
    write-host ("Total Colonnes : ",$nombre_colonnes)

    [string[][]]$tableau = New-Object string[][] ($nombre_lignes +1),($nombre_colonnes +1)

    For ([int]$ligne = $ligne_depart; $ligne -le $nombre_lignes; $ligne++) {
        foreach($colonneK in $global:sharecopter_columns.keys){

            [int] $colonne = $global:sharecopter_columns[$colonneK]

            If ($Worksheet.Cells.Item($ligne,$colonne).Value()){
                $tableau[$ligne - $ligne_depart][$colonne] = $Worksheet.Cells.Item($ligne, $colonne).Value()
            } Else {
                $tableau[$ligne - $ligne_depart][$colonne] = "N/A"
            }
        }
    }

    # Fermeture du fichier et d'Excel
    $excel.Quit()
    $Excel = $null

    # Arrêt de tous les processus Excel en cours d'exécution sur la machine
    Stop-Process -Name EXCEL

    Return($tableau)
}

function Load_File_In_Memory([string] $file_name,[int[]]$colonnes_extraire) {

    [int] $ligne_depart = 2

    $Excel = New-Object -ComObject excel.application 
    $Excel.visible = $False

    # Ouverture du fichier Excel et accès à la feuille numéro 1
    $Workbook = $excel.Workbooks.open($file_name) 
    $Worksheet = $Workbook.WorkSheets.Item(1)
    $Worksheet.activate()

    # Calcul du nombre de lignes
    [int] $nombre_lignes = $Worksheet.UsedRange.Rows.Count
    [int] $nombre_colonnes = $Worksheet.UsedRange.Columns.Count

    # Allocation du tableau dynamique qui contiendra les valeurs
    [int] $total_lignes = ($nombre_lignes - $ligne_depart + 1)

    write-host ("Total Lignes : ",$total_lignes)
    write-host ("Total Colonnes : ",$nombre_colonnes)

    [string[][]]$tableau = New-Object string[][] ($nombre_lignes +1),($nombre_colonnes +1)

    For ([int]$ligne = $ligne_depart; $ligne -le $nombre_lignes; $ligne++) {
        foreach($colonne in $colonnes_extraire){

            If ($Worksheet.Cells.Item($ligne,$colonne)){
                $tableau[$ligne - $ligne_depart][$colonne] = $Worksheet.Cells.Item($ligne, $colonne).Value()
            } Else {
                $tableau[$ligne - $ligne_depart][$colonne] = "N/A"
            }
        }
    }

    # Fermeture du fichier et d'Excel
    $excel.Quit()
    $Excel = $null

    # Arrêt de tous les processus Excel en cours d'exécution sur la machine
    Stop-Process -Name EXCEL

    Return($tableau)
}

######################################################################
# Fonction permettant d'initialiser la liste des colonnes du fichier ShareCopter
# Retour : retourne la hashtable des positions des champs dans le tableau ShareCopter
######################################################################
function Get_ShareCopter_Index_Column {
    [Hashtable]$sharecopter_columns = New-Object Hashtable

    $global:sharecopter_columns = @{
        "ID_CLIENT" = 1
        "STATUS" = 2
        "COMPANY_NAME" = 3
        "A/C NAME" = 4
        "CYCLE" = 5
        "HTAWS-DMAP" = 6
        "HELIONIX STANDARD" = 7
        "SUBSCRIPTION" = 8
        "Delivery method" = 9
        "MAP1" = 10
        "MAP2" = 11
        "MAP3" = 12
        "MAP4" = 13
        "ADRESSE" = 20
        "DB1" = 21
        "DB2" = 22
        "DB3" = 23
        "DB4" = 24
    }
}


######################################################################
# Fonction d'ajout d'un utilisateur
######################################################################
function addUser
{
 param([string]$Key,[string]$Value)

 $d=New-Object Hashtable
 $d | Add-Member -Name Key -MemberType NoteProperty -Value $Key
 $d | Add-Member -Name Value -MemberType NoteProperty -Value $Value

 return $d
}

######################################################################
# Fonction permettant de remplir le créer la hashtable avec les bonnes clés et les valeurs vides
######################################################################
function Get_HashTable
{
    $hashTable = @()
    $ValHashKey=""
    
    for($i = 0; $i -le 8; $i++) 
    {
        $ValHashKey = "AH_TAG_" +[char]([int][char]'A'+$i)
        $value = [char]([int][char]'A'+$i)
        $hashTable+=addUser -Key $ValHashKey -Value $value
    }

    for($i=0; $i -le 3; $i++)
    {
         for($j=0; $j -le 8; $j++)
        {
            $ValHashKey = "AH_TAG_"+ [int]$i +[int]$j
            $value = [int]$i +[int]$j
            $hashTable+=addUser -Key $ValHashKey -Value $value                                  
        }
    }
    return $hashTable
}

######################################################################
# Fonction de génération du fichier Coc
# Entrée : $sharecopter_value --> la liste des champs ShareCopter correspondant à la commande
# Retour : le nom du fichier CoC
######################################################################
function Get_CoC_FileName
{
    param([string[]]$sharecopter_value)

    [string]$file_name = "Coc"
    $id_coc=$sharecopter_value[$sharecopter_columns["ID_CLIENT"]] + [DateTime]::Now.ToString("dd") + [DateTime]::Now.ToString("MM") + [DateTime]::Now.ToString("yyyy")
    [string]$sharecopter_id_client = $sharecopter_value[$sharecopter_columns["ID_CLIENT"]].PadLeft(5, "0")
    [string]$sharecopter_company = $sharecopter_value[$sharecopter_columns["COMPANY_NAME"]]
    [string]$date_jour = [DateTime]::Now.ToString("yyyy") + "_" + [DateTime]::Now.ToString("MM") + "_" + [DateTime]::Now.ToString("dd")

    $file_name += "_" + $id_coc + "_" + $sharecopter_company + "_" + $sharecopter_id_client + "_" + $date_jour +".docx"

    Return ($file_name)
}

######################################################################
# Fonction de création de la hashtable avec les bonnes clés et les valeurs vides pour le remplissage du fichier COC
######################################################################
function Get_HashTable_AH_TAG
{
    [Hashtable]$hashTable = New-Object Hashtable
    [string]$ValHashKey=""
    
    for($i = 0; $i -lt 8; $i++) 
    {
        $ValHashKey = "AH_TAG_" +[char]([int][char]'A'+$i)
        $hashTable.Add($ValHashKey,"")
    }

    for($i=0; $i -le 3; $i++)
    {
         for($j=0; $j -le 8; $j++)
        {
            $ValHashKey = "AH_TAG_"+ [int]$i +[int]$j
            $hashTable.Add($ValHashKey, "")
        }
    }

    Return ($hashTable)
}

######################################################################
# Fonction recherchant dans le tableau des valeurs Export la ligne correspondant à la commande
# Entrée : $reg_expr --> expression régulière
#          $export_contenu --> contenu du fichier export.xlsx
######################################################################
function Find_Export_Line
{
    param([string]$reg_expr,
        [string[][]]$export_contenu)

    ForEach ($export_ligne in $export_contenu)
    {
            #que pour le $reg_
        If ($export_ligne[0] -match $reg_expr)
        {
            Return($export_ligne[0])
        }
    }

    Return($null)
}

######################################################################
# Fonction recherchant récursivement dans le répertoire passé en paramètre tous les fichiers zip non cycliques
# Entrée : $root_directory --> le chemin absolu du répertoire
# Sortie : un tableau de string contenant la liste des fichiers zip
######################################################################
function GetZIPNC
{
    param([string]$root_directory)

    [string[]]$zipnc = @()

    foreach($file in Get-ChildItem -recurse $root_directory -Filter "*.zip*" | Where-Object { $_.FullName -like "*Terrain*"})
    {
        $zipnc += $file.Name
    }

    return $zipnc
}

######################################################################
# Fonction recherchant récursivement dans le répertoire passé en paramètre tous les fichiers zip cycliques
# Entrée : $root_directory --> le chemin absolu du répertoire
# Sortie : un tableau de string contenant la liste des fichiers zip
######################################################################
function GetZIPC
{
    param([string]$root_directory)

    [string[]]$zipnc = @()

    foreach($file in Get-ChildItem -recurse $root_directory -Filter "*.zip*" | Where-Object { $_.FullName -like "*Navigation and Obstacles*"})
    {
        $zipnc += $file.Name
    }

    return $zipnc
}

######################################################################
# Fonction recherchant le fichier zip non cyclique lié
# Entrée : $value --> la référence de la carte / commande
#          $zipnc --> la liste des fichiers zip non cycliques
# Sortie : le nom du fichier zip correspondant / sinon message non trouvé
######################################################################
function SearchZIPNC
{
    param(
        [string] $value,
        [string[]] $zipnc)

    foreach($file in $zipnc)
    {
        If ($file -match $value)
        {
            If ($file -match ".zip.")
            {
                $file = $file.Substring(0, $file.Length - 4)
            }
            Return($file)
        }
    }

    return("Zip non trouvé")
}

######################################################################
# Fonction 
# Entrée : 
#          
#
######################################################################
function Fill_Commande_Map
{
    param([int]$index_map,
        [string[]]$sharecopter_commande,
        [string[][]]$export_contenu,
        [string[]]$zipnc)

    [string]$value = ""
    [string]$reg_expr = ""
    [int]$current_cycle = 0

    #Write-Host $sharecopter_columns["MAP" + $index_map]
    #Write-Host $sharecopter_commande[$sharecopter_columns["MAP" + $index_map]]

    If (-not ($sharecopter_commande[$sharecopter_columns["MAP" + $index_map]] -eq "N/A")){
        # Champs 00 / 10 / 20 / 30
        $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 0)] = $index_map.ToString()

        # Champs 01 / 11 / 21 / 31
        $value01 = $sharecopter_commande[$sharecopter_columns["DB" + $index_map.ToString()]]
        $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 1)] = $value01

        # Champs 02 / 12 / 22 / 32
        $value = $sharecopter_commande[$sharecopter_columns["MAP" + $index_map.ToString()]].Split('-')[1].TrimStart(' ')
        $value += "_" + $sharecopter_commande[$sharecopter_columns["HTAWS-DMAP"]] + "_" + $sharecopter_commande[$sharecopter_columns["HELIONIX STANDARD"]]
        $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 2)] = $value

        # Champs 03 / 13 / 23 / 33
        $reg_expr = $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 1)] + "._00.._.."

        if($reg_expr[4] -eq "M"){
 
            $value1= $reg_expr.Substring(0, 4)
            $value2= $reg_expr.Substring(5, 2)
            
            $reg_expr=$value1+"S"+$value2+ "._00.._.."
 
            $value = Find_Export_Line $reg_expr $export_contenu
            $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 3)] = $value + ".zip"
            write-host ("LA nouvelle valeur de reg apres le if : ",$reg_expr) 
        } else {       
            $value = Find_Export_Line $reg_expr $export_contenu
            $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 3)] = $value + ".zip"
        }

        # Champs 04 / 14 / 24 / 34
        $value = [DateTime]::Now.ToString("MM/dd/yyyy")
        $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 4)] = "Navigation and Obstacles – AIRAC Cycle n°" + $myCycleAIRAC

        # Champs 05 / 15 / 25 / 35
        $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 5)] = "Navigation and Obstacles / " + $sharecopter_commande[$sharecopter_columns["MAP" + $index_map.ToString()]].Split("-")[0].Trim(' ').Substring(0, 4)

        # Champs 06 / 16 / 26 / 36
        # Champs 07 / 17 / 27 / 37
        [int]$cycle=$sharecopter_commande[$sharecopter_columns["CYCLE"]].Split('/')[1]

        # Cas du nouveau client
        if ($myCycleAIRAC -eq $cycle) {
            $reg_expr = $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 1)] + "._..00_.."
            $value = Find_Export_Line $reg_expr $export_contenu

            # Champs X6
            $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 6)] = $value +".zip"
           
            if ($global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 6)] -match ".zip")
            {
                $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 6)] = SearchZIPNC $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 1)] $zipnc
            }

            $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 6)]=$global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 6)] + " Non cycle data"
          
            # Champs X7
            $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 7)] = $terrain + " Non cycle data"
        
            # Champs 08 / 18 / 28 / 38
            $value = $sharecopter_commande[$sharecopter_columns["MAP" + $index_map.ToString()]].Split("-")[1].Trim(' ').Substring(0, 4)
            If ($sharecopter_commande[$sharecopter_columns["HTAWS-DMAP"]].Contains("DMAP")) {
                $value += " / Terrain HTAWS + DMAP"
            } Else {
                $value += " / Terrain HTAWS"
            }

            $global:coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 8)] = $value + " Non cycle data"
        }
    }
}

######################################################################
# Fonction de remplissage d'un fichier CoC
# Entrée : $sharecopter_commande --> la ligne de commande issue du fichier ShareCopter
#          $export_contenu --> tableau du contenu du fichier Excel Export
#          $SignatureFile --> signature du fichier 
# 
######################################################################
function Fill_CoC_File
{
    param([string[]]$sharecopter_commande,
        [string[][]]$export_contenu,
        [string]$SignatureFile,
        [string[]]$zipnc)

    [int]$index_map = 1

    [int]$id_client = $sharecopter_commande[$sharecopter_columns["ID_CLIENT"]]
    [Hashtable]$global:coc_values = Get_HashTable_AH_TAG

    $global:coc_values["AH_TAG_A"] = $sharecopter_commande[$sharecopter_columns["ID_CLIENT"]] + [DateTime]::Now.ToString("dd") + [DateTime]::Now.ToString("MM") + [DateTime]::Now.ToString("yyyy")
    $global:coc_values["AH_TAG_B"] = $sharecopter_commande[$sharecopter_columns["ADRESSE"]]
    $global:coc_values["AH_TAG_C"] = $sharecopter_commande[$sharecopter_columns["ID_CLIENT"]].PadLeft(5, "0")
    $global:coc_values["AH_TAG_D"] = "None"
    $global:coc_values["AH_TAG_E"] = ""
    $global:coc_values["AH_TAG_F"] = [datetime]::Now.ToString("dd/MM/yyyy")
    $global:coc_values["AH_TAG_G"] = "None"
    $global:coc_values["AH_TAG_H"] = ""

    $index_map = 1

    while($index_map -le 4) {
        Fill_Commande_Map $index_map $sharecopter_commande $global:coc_values $export_contenu $zipnc
        $index_map ++
    }

    [string]$nameCoC = Get_CoC_FileName $sharecopter_commande
    readWord $global:coc_values $nameCoC $SignatureFile

    [String]$id_client=$sharecopter_contenu[$i][$sharecopter_columns["ID_CLIENT"]]
    [String]$company_name=$sharecopter_contenu[$i][$sharecopter_columns["COMPANY_NAME"]]
    [String]$required_date=$sharecopter_contenu[$i][$sharecopter_columns["CYCLE"]]
        
    [string] $subscription=""
    If ($subcription_type.Contains("Annual")){
        $subscription="Annual"
    } ElseIf ($subcription_type.Contains("Yearly")){
        $subscription="Annual"
    } Else {
        $subscription="Single"
    }

    [String] $status_generation="ok"
        
    [String] $ligne=$id_client + ';' + $company_name + ';' + $sharecopter_contenu[$i][$sharecopter_columns["STATUS"]] + ';' + $nameCoC.Replace('.docs', '.pdf') + ';' + $required_date + ';' + $subscription + ';' + $status_generation

    Add-Content -Path $FicLogWP2 -Value $ligne

    return $nameCoC
}

function ShowLogs()
{
    $Excel = New-Object -comobject Excel.Application
    $Excel.visible = $true
    $Excel.DisplayAlerts = $False

    $WorkBook = $Excel.Workbooks.open($FicLogWP2) 
}


############################# Programme général ##########################################
############################# !!!!!!!!! Variable d'ENV !!!!!!! ###########################    
# Fichiers input
$COC_TEMPLATE_FILE="$WorkingDirectory\coc_template.docx"
$SHARECOPTERLIST_FILE="$WorkingDirectory\sharecopterList.xlsx"
$EXPORT_FILE="$WorkingDirectory\Export.xlsx"
############################# !!!!!!!!! Variable d'ENV !!!!!!! ###########################    

# Test présence des fichiers 
If (! (Test-Path $global:TemplateCoC))
{
    Write-Host "Le fichier" $global:TemplateCoC "est introuvable"
    Exit -1
}

If (! (Test-Path $SHARECOPTERLIST_FILE))
{
    Write-Host "Le fichier" $SHARECOPTERLIST_FILE "est introuvable"
    Exit -3
}

clear-host

write-host "Lancement du traitement WP2 ..."

write-host "Parsing des colonnes a traiter ..."

Get_ShareCopter_Index_Column

# Chargement en mémoire du contenu du fichier Sharecopter 
write-host "Chargement en mémoire du contenu du fichier Sharecopter  ..."
[string[][]]$sharecopter_contenu = Load_Sharecopter_In_Memory $SHARECOPTERLIST_FILE


# Chargement en mémoire du contenu du fichier Export
Write-Host "Chargement en mémoire du contenu du fichier Export ..."
[string[][]]$export_contenu = Load_File_In_Memory $EXPORT_FILE @(1)

[string[]]$zipnc = GetZIPNC $WP2_GARLABAN_ROOTDIR

$SignatureFile = ChoixSignature

Add-Content -Path $FicLogWP2 -value "Subscription ID;Company name;Status subscription;CoC name generated;required date;Subscription type;status of the COC generation"

#Création d'une boucle pour chaque ligne de facture CoC
for ($i=0;$i -lt $sharecopter_contenu.Length; $i++)
{
    [String]$required_date=$sharecopter_contenu[$i][$sharecopter_columns["CYCLE"]]


    # Prise en compte des lignes CoC dont le status est Not Cancel
    If (($sharecopter_contenu[$i][$sharecopter_columns["STATUS"]] -eq "Resolved") -and ($required_date.Length -gt 0)) {

        [int]$cycle=$required_date.Split('/')[1].TrimStart(' ')

        [String]$subcription_type=$sharecopter_contenu[$i][$sharecopter_columns["SUBSCRIPTION"]]
        [String]$id_client=$sharecopter_contenu[$i][$sharecopter_columns["ID_CLIENT"]]
        [String]$company_name=$sharecopter_contenu[$i][$sharecopter_columns["COMPANY_NAME"]]
        
        [String]$log_coc_name=""
 
        If (($subcription_type.Contains("Annual") -or $subcription_type.Contains("Yearly") -and ($cycle -ge ($myCycleAIRAC - 150)) -and ($cycle -le $myCycleAIRAC)) -or
            ($subcription_type.Contains("Single") -and $cycle -eq $myCycleAIRAC)){
        
            Fill_CoC_File $sharecopter_contenu[$i] $export_contenu $SignatureFile $zipnc
        } Else {

            [string] $subscription=""
            If ($subcription_type.Contains("Annual")){
                $subscription="Annual"
            } ElseIf ($subcription_type.Contains("Yearly")){
                $subscription="Annual"
            } Else {
                $subscription="Single"
            }

            [String] $status_generation=""
            $status_generation="ko"
        
            [String] $ligne=$id_client + ';' + $company_name + ';' + $sharecopter_contenu[$i][$sharecopter_columns["STATUS"]] + ';' + $log_coc_name + ';' + $required_date + ';' + $subscription + ';' + $status_generation

            Add-Content -Path $FicLogWP2 -Value $ligne
        }
    } Else {

        [int]$cycle=$required_date.Split('/')[1].TrimStart(' ')
        [String]$subcription_type=$sharecopter_contenu[$i][$sharecopter_columns["SUBSCRIPTION"]]
        [String]$id_client=$sharecopter_contenu[$i][$sharecopter_columns["ID_CLIENT"]]
        [String]$company_name=$sharecopter_contenu[$i][$sharecopter_columns["COMPANY_NAME"]]
        [String]$log_coc_name=""

        [string] $subscription=""
        If ($subcription_type.Contains("Annual")){
            $subscription="Annual"
        } ElseIf ($subcription_type.Contains("Yearly")){
            $subscription="Annual"
        } Else {
            $subscription="Single"
        }

        [String] $status_generation=""
        $status_generation="ko"
        
        [String] $ligne=$id_client + ';' + $company_name + ';' + $sharecopter_contenu[$i][$sharecopter_columns["STATUS"]] + ';' + $log_coc_name + ';' + $required_date + ';' + $subscription + ';' + $status_generation

        Add-Content -Path $FicLogWP2 -Value $ligne
    }
}

$tab = Get_HashTable

# Conversion des CoC obtenus en PDF
ConvertDocxtoPDF

ShowLogs

######################################################################
# Fin
######################################################################
