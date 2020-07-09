#14/03/2017 14 Mars 2017

$RacineGarlaban ="\\garlaban\fdat1010\Database Production\Generated data\"
$RacineSharecopter ="https://sharecopter-ecm-basic.eurocopter.com/sites/spsa1893/Customer/"
$SourceDir ="\\garlaban\fdat1010\Database Production\Generated data\Temporary Store"
$SourceDirPreviousDB ="\\garlaban\fdat1010\Database Production\Generated data\old DB\"
$CoverageCSVFile="\\garlaban\fdat1010\Database Production\Generated data\Indus\Coverage.csv"
$CycleFile="\\garlaban\fdat1010\Database Production\Generated data\Indus\Cycle.csv"
$CoverageExcelFile="\\garlaban\fdat1010\Database Production\Generated data\Indus\Coverage.xlsx"
$SignaturePath="\\garlaban\fdat1010\Database Production\Generated data\Indus\Signatures"
$GeneratedData="\\garlaban\fdat1010\Database Production\Generated data\"

#Versions des Software MFP
$MFPSW01=@()
$MFPSW02=@()
$MFPSW03=@()
$MFPSW01="D463C01S0400","D463C01S0401"
$MFPSW02="M463C01S0302"
$MFPSW03= "D463C01S0501","M463C02S0510", "E463C01S0510"
$MFPSW03Step2="D463C01S0500","M463C02S0500"

#Versions des DTD
$DTDSW01=@()
$DTDSW02=@()
$DTDSW03=@()
$DTDSW01="X429C40A1003"
$DTDSW02="X429C40A1003"
$DTDSW03="X429C40A1003"

#Templates PDF
$TemplateCyclic = "\\garlaban\fdat1010\Database Production\Generated data\Indus\Templates\TemplateCyclic.docx"
$TemplateNonCyclic2 = "\\garlaban\fdat1010\Database Production\Generated data\Indus\Templates\TemplateNonCyclic2.docx"
$TemplateNonCyclic3 = "\\garlaban\fdat1010\Database Production\Generated data\Indus\Templates\TemplateNonCyclic3.docx"
$iTextPath="\\garlaban\fdat1010\Database Production\Generated data\Indus\itextsharp.dll"

$PCProd1 ="\\ads-pc"
$PCProd2 ="\\ads2-pc"
$PCProd3 ="\\mac17187.eu.eurocopter.corp"
$PCProd4 ="\\mac17213.eu.eurocopter.corp"
$PCProd5 =""
$PCProd6 =""
$PCProd7 =""
$145 = "H,P,T"
$175 = "H,D"
$CoCVierge ="\\garlaban\fdat1010\Database Production\Generated data\CoC Clients"
$CoCDepot =$CocVierge+"\CoC cycle 1502"


################################################# WP4 #############################################################
#CC pour WP4
$ccVAR = "support.ads.ah@airbus.com"

# Dossier dépôt CoC + Extract sharecopter
$SharecoptersVAR = "\\garlaban\fdat1010\Database Production\Generated data\Input_files\sharecopterList.xlsx" 
$contenersVAR = "\\garlaban\fdat1010\Database Production\Generated data\Input_files\"

$SaveExport = "\\garlaban\fdat1010\Database Production\Generated data\Input_files\"
################################################# FIN WP4 #############################################################


################################################# WP2 ###############################
$TemplateCoC =  "\\garlaban\fdat1010\Database Production\Generated data\Indus\Templates\coc_template.docx"

$InputCoC = "\\garlaban\fdat1010\Database Production\Generated data\Input_files\"
$InputCoC2 = "\\garlaban\fdat1010\Database Production\Generated data\Input_files\"
$ExportXlsx = "\\garlaban\fdat1010\Database Production\Generated data\Input_files\Export.xlsx"

$ExportCSV = "\\garlaban\fdat1010\Database Production\Generated data\Input_files\"
 
$WorkingDirectory1 = "\\garlaban\fdat1010\Database Production\Generated data\Input_files\"
$terrain = "TERRAIN"
$WP2_GARLABAN_ROOTDIR = "\\garlaban\fdat1010\Database Production\Generated data\"

$FicLogWP2="\\garlaban\fdat1010\Database Production\Generated data\Input_files\WP2_log.csv"
################################################ FIN WP2 #############################

################################################# WP6 ###############################
