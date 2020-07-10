######################################################################
# Frédéric VERNET
######################################################################

param([Parameter(Mandatory=$true, ValueFromPipeline=$false)]
      [String] $myCycleAIRAC)

Set-PSDebug -Strict

######################################################################
# Récupération des variables d'environnement
######################################################################
$MFPSW01=@()
$MFPSW02=@()

#Crétation de l'alias 7zip
if (test-path "${env:ProgramFiles(x86)}\7-Zip\7z.exe")
{
    set-alias 7z "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
} 
elseif (test-path "${env:ProgramFiles}\7-Zip\7z.exe")
{
    set-alias 7z "${env:ProgramFiles}\7-Zip\7z.exe"
}
else
{
    throw "$env:ProgramFiles\7-Zip\7z.exe needed"
}

#Récupération des variables de paramétrage
$ScriptPath = (Split-Path ((Get-Variable MyInvocation).Value).MyCommand.Path)
$MesVariables = $ScriptPath + "\Variables.ps1"
. $MesVariables

If (Test-Path -Path $FicLogWP2)
{
    Remove-Item $FicLogWP2
}

######################################################################
# Fonction de calcul de cycle AIRAC
# Entrée : $date --> date à laquelle on veut connaître le cycle AIRAC
# Retour : le cycle AIRAC de la forme XXYY
######################################################################
function Get_AIRAC_Cycle
{
    param ([DateTime] $date)

    [DateTime]$date_depart = "01/23/2003"
    [int]$cycle_depart = 1
    [int]$nb_cycles_par_an = 13
    [int]$nb_jours_par_cycle = 28

    [int]$annee_depart = $date_depart.ToString("yy")
    [int]$reste = $null
    [int]$delta_jours = ($date - $date_depart).Days
    [int]$nb_total_cycles = [Math]::DivRem($delta_jours, $nb_jours_par_cycle, [ref]$reste)

    [int]$nb_annees = [Math]::DivRem($nb_total_cycles,$nb_cycles_par_an, [ref]$reste)
    [int]$nb_cycles = $nb_total_cycles - ($nb_annees * $nb_cycles_par_an)

    [int]$annees = $annee_depart + $nb_annees
    [int]$cycle = $cycle_depart + $nb_cycles

    Return ($annees * 100 + $cycle)
}

######################################################################
# Fonction de calcul de cycle AIRAC suivant
# Entrée : $date --> date à laquelle on veut connaître le cycle AIRAC
# Retour : le cycle AIRAC de la forme XXYY
######################################################################
function Get_AIRAC_Cycle_PlusUn
{
    param ([DateTime] $date)

    [DateTime]$date_depart = "01/23/2003"
    [int]$cycle_depart = 1
    [int]$nb_cycles_par_an = 13
    [int]$nb_jours_par_cycle = 28

    [int]$annee_depart = $date_depart.ToString("yy")
    [int]$reste = $null
    [int]$delta_jours = ($date - $date_depart).Days
    [int]$nb_total_cycles = [Math]::DivRem($delta_jours, $nb_jours_par_cycle, [ref]$reste)

    [int]$nb_annees = [Math]::DivRem($nb_total_cycles,$nb_cycles_par_an, [ref]$reste)
    [int]$nb_cycles = $nb_total_cycles - ($nb_annees * $nb_cycles_par_an) + 1

    [int]$annees = $annee_depart + $nb_annees
    [int]$cycle = $cycle_depart + $nb_cycles

    Return ($annees * 100 + $cycle)
}

######################################################################
# Fonction de remplissage du document word
######################################################################
function readWord
{
    param([Hashtable]$coc_values, 
        [string] $name, 
        [string] $SignatureFile,
        [string] $TemplateFile)

    $Word = New-Object -com word.application
    $Word.visible = $False

############################ !!!!!!!!!!! Variable ENV !!!!!!!!!!! ## ok #####################
    $OpenDoc = $Word.documents.Open($TemplateFile)
    $Selection = $Word.selection
    #Debug_Hashtable $coc_values
    $tab = $coc_values
    
    Debug_Hashtable $tab

    $Signature="$SignaturePath\$SignatureFile"
    # Récupération du tableau du template
    $Selection.Tables | ForEach-Object {   
              $objTables = $_      
                                        }

    # Insertion de la bonne signature au bon endroit
    $objTables.Cell(20,1).Range.InlineShapes.AddPicture($Signature)
    
    $coc_values.GetEnumerator() | % {
    #$_["Entry 1"]
    #Modification du template par les bonnes valeurs
        if( $_.Key -eq "AH_TAG_A")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_B")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_C")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_00")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_01")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_02")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_03")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_04")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_05")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_06")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_07")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_08")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_09")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_10")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_11")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_12")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_13")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_14")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_15")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_16")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_17")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_18")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_20")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_21")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_22")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_23")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_24")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_25")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_26")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_27")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_28")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_30")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_31")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_32")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_33")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_34")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_35")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_36")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_37")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_38")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_D")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_E")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_F")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_G")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
        if( $_.Key -eq "AH_TAG_H")
        {
           $Selection.Find.Execute($_.key,$False,$True,$False,$False,$False,$True,1,$False,$_.Value,2)
        }
    }
    
############################# !!!!!!!!! Variable d'ENV !!!!!!! ## ok #########################    
    $chemin = $InputCoC + $name 
    
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
function ConvertDocxtoPDF
{
############################# !!!!!!!!! Variable d'ENV !!!!!!! ## ok #########################    
    $Repertoire = $InputCoC2
    $CoC = Get-ChildItem -recurse $Repertoire

        foreach ($CoCs in $CoC)
        {
            # Pour convertir seulement les docx
            if($CoCs -match "docx") 
            {
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
                #Suppression des CoC qui sont en .docx
                # S'il ne faut pas supprimer les CoC en docx, supprimez les commentaires
                <#if($CoCs -match ".docx")
                {
                    Remove-Item $CoCs.FullName
                }#>
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
function ResearchFile
{
    param ([string] $value)

    [string] $res = ""
    $ExcelWB = new-object -comobject excel.application 
    #Ouverture du fichier Excel Export.xlsx et convertion en fichier excel Export.csv

############################# !!!!!!!!! Variable d'ENV !!!!!!! ## ok #########################    
    $Workbook = $ExcelWB.Workbooks.Open($ExportXlsx) 
    #Ne pas modifier
    $save = $ExportCSV + "Export.csv"
    $Workbook.SaveAs($save,6)
  
    $file = "File"
    # Ne pas changer 
    $Export = Import-Csv -Path $save -Delimiter ";"

    foreach ($Exports in $Export)
    {
        if ($value -match $Exports.$file.substring(0,4) )
        {
            $res = $Exports.$file
        }
    }

    $Workbook.Close($false)
    $ExcelWB.quit()
############################# !!!!!!!!! Variable d'ENV !!!!!!! ## ok #########################    
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
function Load_File_In_Memory
{
    param([string] $file_name,
        [int[]]$colonnes_extraire,
        [int]$ligne_depart,
        [int]$ligne_fin)

    [int]$nombre_lignes = 0
    [int]$nombre_colonnes = 0
    [int]$ligne = 0
    [int]$index_colonne = 0

    $Excel = New-Object -ComObject excel.application 
    $Excel.visible = $False

    # Ouverture du fichier Excel et accès à la feuille numéro 1
    $Workbook = $excel.Workbooks.open($file_name) 
    $Worksheet = $Workbook.WorkSheets.Item(1)
    $Worksheet.activate()

    # Calcul du nombre de lignes
    $nombre_lignes = $Worksheet.UsedRange.Rows.Count
    $nombre_colonnes = $Worksheet.UsedRange.Columns.Count
    $ligne = $ligne_depart

    # Allocation du tableau dynamique qui contiendra les valeurs
    $total_lignes = ($nombre_lignes - $ligne_depart - $ligne_fin + 1)
    $total_colonnes = $colonnes_extraire.Count

    If ($total_lignes -eq 1)
    {
        [string[][]]$tableau = New-Object string[][] ($total_lignes + 1), $total_colonnes
    }
    Else
    {
        [string[][]]$tableau = New-Object string[][] $total_lignes, $total_colonnes
    }

    While ($ligne -le ($nombre_lignes - $ligne_fin))
    {
        $index_colonne = 0

        ForEach($colonne in $colonnes_extraire)
        {
            If ($colonne > $nombre_colonnes)
            {
                Write-Host "Colonne hors des limites du fichier Excel"
            }

            If ($worksheet.Cells.Item($ligne,$colonne))
            {
                $tableau[($ligne - $ligne_depart)][$index_colonne] = $worksheet.Cells.Item($ligne,$colonne).Value()
            }
            Else
            {
                $tableau[($ligne - $ligne_depart)][$index_colonne] = "N/A"
            }

            $index_colonne ++
        }

        $ligne ++
    }

    # Fermeture du fichier et d'Excel
    $excel.Quit()
    $Excel = $null

    # Arrêt de tous les processus Excel en cours d'exécution sur la machine
    Stop-Process -Name EXCEL

    Return($tableau)
}

######################################################################
# Fonction permettant de lister le contenu d'un tableau
# Entrée : $tableau --> le tableau à 2 dimensions
######################################################################
function Debug_Array
{
    param([string[][]] $tableau)

    [int]$index_ligne = 0
    [int]$index_colonne = 0

    while ($index_ligne -lt $tableau.Count)
    {
        $index_colonne = 0

        while ($index_colonne -lt $tableau[$index_ligne].Count)
        {
            Write-Host "Cellule[$index_ligne,$index_colonne] = "$tableau[$index_ligne][$index_colonne]

            $index_colonne ++;
        }

        $index_ligne ++
    }
}

######################################################################
# Fonction permettant de lister le contenu d'une hashtable
# Entrée : $hashtable --> la hashtable
######################################################################
function Debug_Hashtable
{
    param([Hashtable]$hashtable)

    [string]$key = ""

    for($i = 0; $i -lt 8; $i++) 
    {
        $key = "AH_TAG_" +[char]([int][char]'A'+$i)
        Write-Host $key " --> " $hashtable[$key]
    }

    for($i=0; $i -le 3; $i++)
    {
         for($j=0; $j -le 8; $j++)
        {
            $key = "AH_TAG_"+ [int]$i +[int]$j
            Write-Host $key " --> " $hashtable[$key]
        }
    }
}

######################################################################
# Fonction permettant d'initialiser la liste des colonnes du fichier ShareCopter
# Retour : retourne la hashtable des positions des champs dans le tableau ShareCopter
######################################################################
function Get_ShareCopter_Index_Column
{
    [Hashtable]$sharecopter_columns = New-Object Hashtable

    $index = 0

    $sharecopter_columns.Add("ID_CLIENT", $index ++)
    $sharecopter_columns.Add("STATUS", $index ++)
    $sharecopter_columns.Add("COMPANY_NAME", $index ++)
    $sharecopter_columns.Add("A/C NAME", $index ++)
    $sharecopter_columns.Add("CYCLE", $index ++) # Colonne J
    $sharecopter_columns.Add("HTAWS-DMAP", $index ++)
    $sharecopter_columns.Add("HELIONIX STANDARD", $index ++)
    $sharecopter_columns.Add("SUBSCRIPTION", $index ++)
    $sharecopter_columns.Add("Delivery method", $index ++)
    $sharecopter_columns.Add("MAP1", $index ++)
    $sharecopter_columns.Add("MAP2", $index ++)
    $sharecopter_columns.Add("MAP3", $index ++)
    $sharecopter_columns.Add("MAP4", $index ++)
    $sharecopter_columns.Add("ADRESSE", $index ++)
    $sharecopter_columns.Add("DB1", $index ++)
    $sharecopter_columns.Add("DB2", $index ++)
    $sharecopter_columns.Add("DB3", $index ++)
    $sharecopter_columns.Add("DB4", $index ++)

    Return($sharecopter_columns)
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

    [Hashtable]$sharecopter_columns = Get_ShareCopter_Index_Column

    [string]$file_name = "Coc"
    $id_coc=$sharecopter_value[$sharecopter_columns["ID_CLIENT"]] + [DateTime]::Now.ToString("dd") + [DateTime]::Now.ToString("MM") + [DateTime]::Now.ToString("yyyy")
    [string]$sharecopter_id_client = $sharecopter_value[$sharecopter_columns["ID_CLIENT"]].PadLeft(5, "0")
    [string]$sharecopter_company = $sharecopter_value[$sharecopter_columns["COMPANY_NAME"]]
    [string]$date_jour = [DateTime]::Now.ToString("yyyy") + "_" + [DateTime]::Now.ToString("MM") + "_" + [DateTime]::Now.ToString("dd")

    $file_name += "_" + $id_coc + "_" + $sharecopter_company + "_" + $sharecopter_id_client + "_" + $date_jour +".docx"

    Return ($file_name)
}

######################################################################
# Fonction de génération du fichier Coc pour la FAL
# Entrée : $sharecopter_value --> la liste des champs ShareCopter correspondant à la commande
# Retour : le nom du fichier CoC
######################################################################
function Get_CoC_FileName_FAL
{
    param([string]$site, [int] $id_coc)

    [string]$file_name = "Coc_FAL - " + $site + "_00"
    [string]$date_jour = [DateTime]::Now.ToString("yyyy") + "_" + [DateTime]::Now.ToString("MM") + "_" + [DateTime]::Now.ToString("dd")

    $file_name += $id_coc.ToString() + "_" + $date_jour +".docx"

    Return ($file_name)
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
        [Hashtable]$coc_values,
        [string[][]]$export_contenu,
        [string[]]$zipnc)

    [Hashtable]$sharecopter_columns = Get_ShareCopter_Index_Column
    [string]$value = ""
    [string]$reg_expr = ""
    [int]$current_cycle = 0

    If ($sharecopter_commande[$sharecopter_columns["MAP" + $index_map]])
    {
        # Champs 00 / 10 / 20 / 30
        $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 0)] = $index_map.ToString()

        # Champs 01 / 11 / 21 / 31
        $value01 = $sharecopter_commande[$sharecopter_columns["DB" + $index_map.ToString()]]
        $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 1)] = $value01

        # Champs 02 / 12 / 22 / 32
        $value = $sharecopter_commande[$sharecopter_columns["MAP" + $index_map.ToString()]].Split('-')[1].TrimStart(' ')
        $value += "_" + $sharecopter_commande[$sharecopter_columns["HTAWS-DMAP"]] + "_" + $sharecopter_commande[$sharecopter_columns["HELIONIX STANDARD"]]
        $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 2)] = $value

        # Champs 03 / 13 / 23 / 33
        $reg_expr = $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 1)] + "._00.._.."
        $value = Find_Export_Line $reg_expr $export_contenu
        $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 3)] = $value + ".zip"

        # Champs 04 / 14 / 24 / 34
        $value = [DateTime]::Now.ToString("MM/dd/yyyy")
        $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 4)] = "Navigation and Obstacles – AIRAC Cycle n°" + $myCycleAIRAC

        # Champs 05 / 15 / 25 / 35
        $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 5)] = "Navigation and Obstacles / " + $sharecopter_commande[$sharecopter_columns["MAP" + $index_map.ToString()]].Split("-")[0].Trim(' ').Substring(0, 4)

        # Champs 06 / 16 / 26 / 36
        # Champs 07 / 17 / 27 / 37

        [int]$cycle=$sharecopter_commande[$sharecopter_columns["CYCLE"]].Split('/')[1]

        # Cas du nouveau client
        if ($myCycleAIRAC -eq $cycle) 
        {
            $reg_expr = $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 1)] + "._..00_.."
            $value = Find_Export_Line $reg_expr $export_contenu

            # Champs X6
           $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 6)] = $value +".zip"
           
           if ($coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 6)] -match ".zip")
           {
                $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 6)] = SearchZIPNC $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 1)] $zipnc
           }

           $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 6)]=$coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 6)] + " Non cycle data"
          
            # Champs X7
            $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 7)] = $terrain + " Non cycle data"
        
            # Champs 08 / 18 / 28 / 38
            $value = $sharecopter_commande[$sharecopter_columns["MAP" + $index_map.ToString()]].Split("-")[1].Trim(' ').Substring(0, 4)
            If ($sharecopter_commande[$sharecopter_columns["HTAWS-DMAP"]].Contains("DMAP"))
            {
                $value += " / Terrain HTAWS + DMAP"
            }
            Else
            {
                $value += " / Terrain HTAWS"
            }

            $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 8)] = $value + " Non cycle data"
        }
    }
}

######################################################################
# Fonction 
# Entrée : 
#          
#
######################################################################
Function GetFileFALZIP
{
    param([string] $directory)

    $zipFile = Get-ChildItem $directory -Filter "*.zip"

    return $zipFile.FullName
}


######################################################################
# Fonction 
# Entrée : 
#          
#
######################################################################
function Fill_Commande_Map_FAL
{
    param([string] $site,
        [int]$index_map,
        [Hashtable]$coc_values,
        [string[][]]$export_contenu,
        [string[]]$zipnc)

    [string]$value = ""
    [string]$reg_expr = ""
    [int]$current_cycle = 0
    [int]$cycle = 0

    $value = [DateTime]::Now.ToString("MM/dd/yyyy")
    
    [string]$cycleStr = $myCycleAIRAC
    [string] $yearStrCycle = $cycleStr.Substring(0, 2)
    [string] $strCycle = $cycleStr.Substring(2, 2)
    [int] $index = 0

    If ($site -match "MAR")
    {
        $zipFile = GetFileFALZIP $DirectoryFALMAR
    }
    Else
    {
        $zipFile = GetFileFALZIP $DirectoryFALDON
    }

    7z e $zipFile -aoa

    foreach($archiveZip in 7z l $zipFile)
    {
        $index = $archiveZip.LastIndexof("\")

        If ($index -gt 0)
        {
            $file = $archiveZip.Substring($index + 1)

            If ($file -match ".pdf")
            {
                # Champs 03 
                $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 3)] = $file

                # Champ 04
                If ($site -match "MAR")
                {
                    $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 4)] = "France3_HTAWS_DMAP _Std02_Nav_Obs_AIRAC_20" + $yearStrCycle + "_" + $strCycle + "_Delivery note"

                    Copy-Item $file -Destination $GDAT_FAL_MAR\X464C81S0103\X464C81S0103 -Force
                    Move-Item $file -Destination $GDAT_FAL_MAR\X464C82S0103\X464C82S0103 -Force
                }
                Else
                {
                    $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 4)] = "Germany_HTAWS_Std01_Nav_Obs_AIRAC_20"  + $yearStrCycle + "_" + $strCycle + "_Delivery note"
                    Copy-Item $file -Destination $GDAT_FAL_DON\X464C81S0101\X464C81S0101 -Force
                    Move-Item $file -Destination $GDAT_FAL_DON\X464C82S0101\X464C82S0101 -Force
                }
            }

            If ($file -match ".crc")
            {
                If ($site -match "MAR")
                {
                    Copy-Item $file -Destination $GDAT_FAL_MAR\X464C81S0103\X464C81S0103 -Force
                    Move-Item $file -Destination $GDAT_FAL_MAR\X464C82S0103\X464C82S0103 -Force
                }
                Else
                {
                    Copy-Item $file -Destination $GDAT_FAL_DON\X464C81S0101\X464C81S0101 -Force
                    Move-Item $file -Destination $GDAT_FAL_DON\X464C82S0101\X464C82S0101 -Force
                }
            }

            If ($file -match "CAE1")
            {
                # Champs 06 
                $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 6)] = $file

                # Champ 07
                If ($site -match "MAR")
                {
                    $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 7)] = "France3_HTAWS_DMAP _Std02_Nav_Obs_AIRAC_20" + $yearStrCycle + "_" + $strCycle + "_Navigation_set1"
                    Move-Item $file -Destination $GDAT_FAL_MAR\X464C81S0103\X464C81S0103 -Force
                }
                Else
                {
                    $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 7)] = "Germany_HTAWS_Std01_Nav_Obs_AIRAC_20" + $yearStrCycle + "_" + $strCycle + "_Navigation_set1"
                    Move-Item $file -Destination $GDAT_FAL_DON\X464C81S0101\X464C81S0101 -Force
                }
            }

            If ($file -match "CAE2")
            {
                # Champs 09 
                $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 9)] = $file

                # Champ 10
                If ($site -match "MAR")
                {
                    $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 10)] = "France3_HTAWS_DMAP _Std02_Nav_Obs_AIRAC_20"  + $yearStrCycle + "_" + $strCycle +  "_Navigation_set2"
                    Move-Item $file -Destination $GDAT_FAL_MAR\X464C81S0103\X464C81S0103 -Force
                }
                Else
                {
                    $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map - 1) * 10) + 10)] = "Germany_HTAWS_Std01_Nav_Obs_AIRAC_20"  + $yearStrCycle + "_" + $strCycle + "_Navigation_set2"
                    Move-Item $file -Destination $GDAT_FAL_DON\X464C81S0101\X464C81S0101 -Force
                }
            }

            If ($file -match "CBST")
            {
                # Champs 12 
                $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map) * 10) + 2)] = $file

                # Champ 13
                If ($site -match "MAR")
                {
                    $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map) * 10) + 3)] = "France3_HTAWS_DMAP _Std02_Nav_Obs_AIRAC_20" + $yearStrCycle + "_" + $strCycle + "_Obstacles"
                    Move-Item $file -Destination $GDAT_FAL_MAR\X464C82S0103\X464C82S0103 -Force
                }
                Else
                {
                    $coc_values["AH_TAG_" + "{0:D2}" -f ((($index_map) * 10) + 3)] = "Germany_HTAWS_Std01_Nav_Obs_AIRAC_20" + $yearStrCycle + "_" + $strCycle + "_Obstacles"
                    Move-Item $file -Destination $GDAT_FAL_DON\X464C82S0101\X464C82S0101 -Force
                }
            }
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

    [Hashtable]$sharecopter_columns = Get_ShareCopter_Index_Column
    [int]$index_map = 1

    [int]$id_client = $sharecopter_commande[$sharecopter_columns["ID_CLIENT"]]
    [Hashtable]$coc_values = Get_HashTable_AH_TAG

    $coc_values["AH_TAG_A"] = $sharecopter_commande[$sharecopter_columns["ID_CLIENT"]] + [DateTime]::Now.ToString("dd") + [DateTime]::Now.ToString("MM") + [DateTime]::Now.ToString("yyyy")
    $coc_values["AH_TAG_B"] = $sharecopter_commande[$sharecopter_columns["ADRESSE"]]
    $coc_values["AH_TAG_C"] = $sharecopter_commande[$sharecopter_columns["ID_CLIENT"]].PadLeft(5, "0")
    $coc_values["AH_TAG_D"] = "None"
    $coc_values["AH_TAG_E"] = ""
    $coc_values["AH_TAG_F"] = [datetime]::Now.ToString("dd/MM/yyyy")
    $coc_values["AH_TAG_G"] = "None"
    $coc_values["AH_TAG_H"] = ""

    $index_map = 1

    while($index_map -le 4)
    {
        Fill_Commande_Map $index_map $sharecopter_commande $coc_values $export_contenu $zipnc
        write-host ("Index_map = ", $index_map)
        $index_map ++
    }

    [string]$nameCoC = Get_CoC_FileName $sharecopter_commande
    readWord $coc_values $nameCoC $SignatureFile $TemplateCoC

    [String]$id_client=$sharecopter_contenu[$i][$sharecopter_columns["ID_CLIENT"]]
    [String]$company_name=$sharecopter_contenu[$i][$sharecopter_columns["COMPANY_NAME"]]
    [String]$required_date=$sharecopter_contenu[$i][$sharecopter_columns["CYCLE"]]
        
    [string] $subscription=""
    If ($subcription_type.Contains("Annual"))
    {
        $subscription="Annual"
    }
    Else
    {
        $subscription="Single"
    }

    [String] $status_generation="ok"
        
    [String] $ligne=$id_client + ';' + $company_name + ';' + $sharecopter_contenu[$i][$sharecopter_columns["STATUS"]] + ';' + $nameCoC.Replace('.docs', '.pdf') + ';' + $required_date + ';' + $subscription + ';' + $status_generation

    Add-Content -Path $FicLogWP2 -Value $ligne

    return $nameCoC
}

######################################################################
# Fonction de remplissage d'un fichier CoC pour la FAL (France et Allemagne)
# Entrée : $site --> site où la FAL doit être générée
#          $export_contenu --> contenu du fichier Export.xlsx
#          $signatureFile --> fichier signature
#          $zipnc --> fichier zip pour les données non cycliques
######################################################################
function Fill_CoC_File_FAL
{
    param([string] $site,
        [string[][]]$export_contenu,
        [string]$SignatureFile,
        [string[]]$zipnc,
        [int]$id_coc)

    [int]$index_map = 1

    [Hashtable]$coc_values = Get_HashTable_AH_TAG

    $coc_values["AH_TAG_D"] = "None"
    $coc_values["AH_TAG_E"] = ""
    $coc_values["AH_TAG_F"] = [datetime]::Now.ToString("dd/MM/yyyy")
    $coc_values["AH_TAG_G"] = "None"
    $coc_values["AH_TAG_H"] = ""

    $index_map = 1

    Fill_Commande_Map_FAL $site $index_map $coc_values $export_contenu $zipnc

    [string]$nameCoC = ""
        
    If ($site -match "MAR")
    {
        $nameCoC = Get_CoC_FileName_FAL $site 69 + [DateTime]::Now.ToString("dd") + [DateTime]::Now.ToString("MM") + [DateTime]::Now.ToString("yyyy")
        $coc_values["AH_TAG_A"] = "69" + [DateTime]::Now.ToString("dd") + [DateTime]::Now.ToString("MM") + [DateTime]::Now.ToString("yyyy")
        readWord $coc_values $nameCoC $SignatureFile $TemplateCoCFALMAR
    }
    Else
    {
        $nameCoC = Get_CoC_FileName_FAL $site 70 + [DateTime]::Now.ToString("dd") + [DateTime]::Now.ToString("MM") + [DateTime]::Now.ToString("yyyy")
        $coc_values["AH_TAG_A"] = "70" + [DateTime]::Now.ToString("dd") + [DateTime]::Now.ToString("MM") + [DateTime]::Now.ToString("yyyy")
        readWord $coc_values $nameCoC $SignatureFile $TemplateCoCFALDON
    }

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
# Variables globales
[string]$WorkingDirectory = $WorkingDirectory1

# Fichiers input
$COC_TEMPLATE_FILE="$WorkingDirectory\coc_template.docx"
$SHARECOPTERLIST_FILE="$WorkingDirectory\sharecopterList.xlsx"
$EXPORT_FILE="$WorkingDirectory\Export.xlsx"
############################# !!!!!!!!! Variable d'ENV !!!!!!! ###########################    

# Test présence des fichiers 
If (! (Test-Path $TemplateCoC))
{
    Write-Host "Le fichier" $TemplateCoC "est introuvable"
    Exit -1
}

If (! (Test-Path $TemplateCoCFALDON))
{
    Write-Host "Le fichier" $TemplateCoCFALDON "est introuvable"
    Exit -1
}

If (! (Test-Path $TemplateCoCFALMAR))
{
    Write-Host "Le fichier" $TemplateCoCFALMAR "est introuvable"
    Exit -1
}

If (! (Test-Path $SHARECOPTERLIST_FILE))
{
    Write-Host "Le fichier" $SHARECOPTERLIST_FILE "est introuvable"
    Exit -3
}

clear-host
Write-Host "Chargement des fichiers ..."

# Chargement en mémoire du contenu du fichier Sharecopter 
[string[][]]$sharecopter_contenu = Load_File_In_Memory $SHARECOPTERLIST_FILE @(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 20, 21, 22, 23, 24) 2 0
write-host "..."

# Chargement en mémoire du contenu du fichier Export
[string[][]]$export_contenu = Load_File_In_Memory $EXPORT_FILE @(1) 2 0
Write-Host "..."

[string[]]$zipnc = GetZIPNC $WP2_GARLABAN_ROOTDIR

$SignatureFile = ChoixSignature
write-host("Taille max i = ",$sharecopter_contenu.Length)

[Hashtable]$sharecopter_columns = Get_ShareCopter_Index_Column

Add-Content -Path $FicLogWP2 -value "Subscription ID;Company name;Status subscription;CoC name generated;required date;Subscription type;status of the COC generation"

#Création d'une boucle pour chaque ligne de facture CoC
for ($i=0;$i -lt $sharecopter_contenu.Length; $i++)
{
    # Prise en compte des lignes CoC dont le status est Not Cancel
    If (($sharecopter_contenu[$i][$sharecopter_columns["STATUS"]] -eq "Resolved") -and ($sharecopter_contenu[$i][$sharecopter_columns["CYCLE"]].Length -gt 0))
    {
        [String]$subcription_type=$sharecopter_contenu[$i][$sharecopter_columns["SUBSCRIPTION"]]
        [int]$cycle=$sharecopter_contenu[$i][$sharecopter_columns["CYCLE"]].Split('/')[1].TrimStart(' ')
        
        [String]$id_client=$sharecopter_contenu[$i][$sharecopter_columns["ID_CLIENT"]]
        [String]$company_name=$sharecopter_contenu[$i][$sharecopter_columns["COMPANY_NAME"]]
        [String]$required_date=$sharecopter_contenu[$i][$sharecopter_columns["CYCLE"]]
        [String]$log_coc_name=""
        


        If (($subcription_type.Contains("Annual") -and ($cycle -ge ($myCycleAIRAC - 100)) -and ($cycle -le $myCycleAIRAC)) -or
            ($subcription_type.Contains("Single") -and $cycle -eq $myCycleAIRAC))
        {
            Fill_CoC_File $sharecopter_contenu[$i] $export_contenu $SignatureFile $zipnc
        }
        Else
        {
            [string] $subscription=""
            If ($subcription_type.Contains("Annual"))
            {
                $subscription="Annual"
            }
            Else
            {
                $subscription="Single"
            }

            [String] $status_generation=""
            $status_generation="ko"
        
            [String] $ligne=$id_client + ';' + $company_name + ';' + $sharecopter_contenu[$i][$sharecopter_columns["STATUS"]] + ';' + $log_coc_name + ';' + $required_date + ';' + $subscription + ';' + $status_generation

            Add-Content -Path $FicLogWP2 -Value $ligne
        }
    }
    Else
    {
        [String]$subcription_type=$sharecopter_contenu[$i][$sharecopter_columns["SUBSCRIPTION"]]
        [int]$cycle=$sharecopter_contenu[$i][$sharecopter_columns["CYCLE"]].Split('/')[1].TrimStart(' ')
        
        [String]$id_client=$sharecopter_contenu[$i][$sharecopter_columns["ID_CLIENT"]]
        [String]$company_name=$sharecopter_contenu[$i][$sharecopter_columns["COMPANY_NAME"]]
        [String]$required_date=$sharecopter_contenu[$i][$sharecopter_columns["CYCLE"]]
        [String]$log_coc_name=""

        [string] $subscription=""
        If ($subcription_type.Contains("Annual"))
        {
            $subscription="Annual"
        }
        Else
        {
            $subscription="Single"
        }

        [String] $status_generation=""
        $status_generation="ko"
        
        [String] $ligne=$id_client + ';' + $company_name + ';' + $sharecopter_contenu[$i][$sharecopter_columns["STATUS"]] + ';' + $log_coc_name + ';' + $required_date + ';' + $subscription + ';' + $status_generation

        Add-Content -Path $FicLogWP2 -Value $ligne
    }
}
$tab = Get_HashTable

#[string] $falexecbool = $(read-host "Voulez vous produire et delivrer les COC FAL ? oui/non")

If($falexecbool -eq "oui"){

    # Traitement spécifique pour la FAL
    Fill_CoC_File_FAL "MAR" $export_contenu $SignatureFile $zipnc
    Write-Host "Veuillez vérifier le COC de la FAL MAR"
    Pause
    Fill_CoC_File_FAL "DON" $export_contenu $SignatureFile $zipnc
    Write-Host "Veuillez vérifier le COC de la FAL DON"
    Pause
}
    # Conversion des CoC obtenus en PDF
    ConvertDocxtoPDF

If($falexecbool -eq "oui"){
    Copy-Item $InputCoC\*FAL*DON*.pdf $GDAT_FAL_DON -Force
    Copy-Item $InputCoC\*FAL*MAR*.pdf $GDAT_FAL_MAR -Force
}

[string] $mails = $(Read-Host "Voulez-vous générer les mails ? oui/non")
If ($mails -eq "oui")
{
    C:\Users\ADSADM\Desktop\Indus\WP4.ps1
}

ShowLogs

######################################################################
# Fin
######################################################################
