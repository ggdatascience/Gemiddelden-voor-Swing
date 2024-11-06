### Script voor het berekenen van gemiddelden per variabele
#
# De basis van het script is gemaakt door Pieter Seinen 
# Het script werkt met behulp van een configuratie Excel bestand. Hierin staan de volgende opties (zie Voorbeeld_conf.xlsx):
# - bestandsnaam	[volledig/path/naar/spss_bestand_monitor_inclusief_labels.sav]
# - output [volledig/path/naar/output_map]	
# - variabele	[gmj_alcohollft,gmj_suiker_gemiddeld,gmj_energie_gemiddeld] (let op de komma's!)
# - variabele_label	[	Gemiddelde startleeftijd alcohol gebruik,Gemiddeld aantal glazen drankjes met suiker per week,Gemiddeld aantal blikjes energiedrank per week] (let op de komma's!)
# - bron [Gezondheidsmonitor Jeugd]	
# - jaarvariabele	[Jaar] 
# - jaren_voor_analyse [2023]	
# - type_periode [Jaar]	
# - weegfactor [weegfactor_gemeentelijk]	
# - gebiedsniveau	[gemeente] 
# - gebiedsindeling	[Gemeentecode2023]
# - minimum_observaties	[30]
# - crossings	[gmj_opleidingsniveau
#             gmj_geslacht
#             gmj_klas] (in een lijst onder elkaar)
# - namen_crossings	[Onderwijssoort in twee groepen
#                    Gender in twee categorieën
#                    Leerjaar in twee categorieën] (in een lijst onder elkaar)
# - labels_name	[Vmbo,Havo/Vwo
#               Jongen,Meisje
#               Klas 2,Klas 4] (in een lijst onder elkaar)
# - labels_id [1, 2
#              1, 2
#              1, 2] (in een lijst onder elkaar)

#
# Dit script voegt de gemeenten automatisch samen tot regio en zet het GGD gemiddelde in de Excel file erbij.
# De crossings worden los uitgedraaid per indicator.
#
# GGD Zuid-Limburg
# 
# Updates
# Update 27 mei 2024: Gemiddelde over de GGD regio en provincie te berekenen; gebruik weegfactor_landelijk.
# Update 17 juni 2024: Weegfactor op 1 zetten als weegfactor ontbrekend is.
# Update 6 november 2024: Uitleg en install.packages toegevoegd 
#
# Libraries inladen - indien nodig kunnen de libraries geinstalleerd worden met install.packages.
install.packages("haven","labelled","openxlsx","dplyr","glue","kableExtra","stringr","tidyr")
library(haven)
library(labelled)
library(openxlsx)
library(dplyr)
library(glue)
library(kableExtra)
library(stringr)
library(tidyr)
# Gebruik CTRL+ALT+E om het script vanaf hier te draaien
rm(list = ls())

# Lees het configuratiebestand in
config_gem <- openxlsx::read.xlsx("/volledig/path/naar/configuratie_excel_bestand.xlsx")

# Omzetten naar variabelen die gebruikt kunnen worden
bestandsnaam <- config_gem$bestandsnaam[1]
basismap_output <- config_gem$output[1]
vars_list <- unlist((str_split(config_gem$variabele, ",")))
vars_list <- vars_list[!is.na(vars_list)]
vars_labels <- unlist((str_split(config_gem$variabele_label, ",")))
vars_labels <- vars_labels[!is.na(vars_labels)]
bron <- config_gem$bron[1]
jaren_voor_analyse <- config_gem$jaren_voor_analyse[1]
jaarvariabele <- as.character(config_gem$jaren_voor_analyse[1])
type_periode <- config_gem$jaarvariabele[1]  
weegfactor <- config_gem$weegfactor[1]
gebiedsniveau <- config_gem$gebiedsniveau[1]
gebiedsindeling <- config_gem$gebiedsindeling[1]
min_observaties <- config_gem$minimum_observaties[1]
crossings_list <- unlist(config_gem$crossings[!is.na(config_gem$crossings)])
namen_crossings <- unlist(config_gem$namen_crossings[!is.na(config_gem$namen_crossings)])
variabel_labels <- "gem"
basismap_output="C:/Users/stijn.michielse/Documents"

# Convenant stelt dat 'microdata' niet gedeeld mag worden met 3en. 
# GGDGHOR verstaat daaronder ook groepsindelingen van 1.
# Minimum aantal observaties per groepsindeling waarbij data geupload mag worden naar ABF 
minimum_obs_per_rij <- 5
# Waarde -99996 past bij Swings default Special value voor 'missing'
missing_voor_privacy <- -99996
nr_regio <- 1
data_totaal <-  haven::read_spss(paste(bestandsnaam, sep = ""))

# Dataframe voorbereiden
jaren_voor_analyse <- str_split(jaren_voor_analyse, ",") %>% unlist() %>% as.numeric()
data_totaal[[jaarvariabele]] <- jaren_voor_analyse
data_totaal <- data_totaal %>%  filter(Onderzoeksjaar == c(jaarvariabele))

# Wanneer weegfactor ontbreekt, deze op 1 zetten, anders waarden overnemen
if (is.na(weegfactor)){
  data_totaal$weegfactor = 1
} else {
  data_totaal$weegfactor <- data_totaal[[weegfactor]]
}

l <- 0
#Voor alle variabelen 
for(variabele in vars_list) {
k <- 0
  # Voor alle crossings
  for (crossings in crossings_list) {
    k <- k+1
    if (gebiedsniveau == "gemeente") {
  # Selecteren vanuit DF naar kubus en gewogen getallen produceren.
  kubus_df <- data_totaal %>%
  filter(jaarvariabele %in% jaren_voor_analyse,
         !is.na(.[[variabele]]),
         !is.na(.[[gebiedsindeling]]),
         !is.na(weegfactor),
         across(.cols = all_of(crossings) ,
                .fns = ~ !is.na(.x))) %>%
  mutate(n = 1,
         var = as.numeric(.[[variabele]],
                      levels = val_labels(.[[variabele]]),
                      labels = names(val_labels(.[[variabele]]))))%>%
  group_by(.[[jaarvariabele]],.[[gebiedsindeling]], across(all_of(crossings)), var) %>%
  summarise(n_gewogen = sum(weegfactor),
            n_ongewogen = sum(n)) %>%
  ungroup()

#Functie om te kleine aantallen per groep te verwijderen 
verwijder_kleine_aantallen <- function(x, ongewogen){if(ongewogen < minimum_obs_per_rij & ongewogen != missing_voor_privacy){missing_voor_privacy}else{x}}

# Bereken het gewogen gemiddelde per gebiedsniveau en per crossing, afronden op 1 cijfer achter de komma
kubus_df <- kubus_df %>%
  mutate(var_gewogen = var * n_gewogen) %>%
  mutate(var_ongewogen = var * n_ongewogen) %>%
  group_by(`.[[gebiedsindeling]]`, across(all_of(crossings))) %>%
  summarise(Gew_ge=round(sum(var_gewogen)/sum(n_gewogen), digits = 1),
            n_ongewogen=sum(n_ongewogen)) %>%
  #Verwijder te lage aantallen
  mutate(across(.cols = "n_ongewogen" ,
                .fns = Vectorize(verwijder_kleine_aantallen), ongewogen = n_ongewogen)) %>%
  ungroup()
kubus_df1 <<- kubus_df

# Vang nullen als er geen observaties zijn
kubus_df[4][kubus_df[4] == 0] <- missing_voor_privacy
kubus_df[3][kubus_df[4] == 0] <- missing_voor_privacy

# Inbouwen vangnet minimum observaties per cel. Expliciet gemaakt!
kubus_df <- kubus_df %>% mutate(kubus_df, n_ongewogen = case_when(n_ongewogen < minimum_obs_per_rij ~ missing_voor_privacy, TRUE ~ n_ongewogen))
kubus_df <- kubus_df %>% mutate(kubus_df, Gew_ge = case_when(n_ongewogen < minimum_obs_per_rij ~ missing_voor_privacy, TRUE ~ Gew_ge))


# Voeg kolom toe met geolevel op basis van gebiedsniveau
kubus_df <- cbind(Geolevel = gebiedsniveau, kubus_df)
# Voeg kolom toe met jaarvariabele
kubus_df <- cbind(Jaar = jaarvariabele, kubus_df)

#Kolomnamen toewijzen 
names(kubus_df) <- c("Jaar","Geolevel","Geoitem",crossings,glue("{variabele}_{variabel_labels}"),'n_ongewogen')

## Toevoeging om het gemiddelde over de GGD regio te berekenen.
# Totaal voor GGD regio uitgesplitst en geheel. 27 mei 2024: weegfactor_landelijk toegevoegd
kubus_df_tot <- data_totaal %>%
  filter(jaarvariabele %in% jaren_voor_analyse,
         !is.na(.[[variabele]]),
         !is.na(.[[gebiedsindeling]]),
         !is.na(weegfactor_landelijk),
         across(.cols = all_of(crossings) ,
                .fns = ~ !is.na(.x))) %>%
  mutate(n = 1,
         var = as.numeric(.[[variabele]],
                          levels = val_labels(.[[variabele]]),
                          labels = names(val_labels(.[[variabele]]))))%>%
  group_by(.[[jaarvariabele]], across(all_of(crossings)), var) %>%
  summarise(n_gewogen = sum(weegfactor_landelijk),
            n_ongewogen = sum(n)) %>%
  ungroup()

# Bereken het gewogen gemiddelde voor ZL, afronden op 2 cijfers achter de komma
kubus_df_tot <- kubus_df_tot %>%
  mutate(var_gewogen = var * n_gewogen) %>%
  mutate(var_ongewogen = var * n_ongewogen) %>%
  group_by(across(all_of(crossings))) %>%
  summarise(Gew_ge=round(sum(var_gewogen)/sum(n_gewogen), digits = 2),
            n_ongewogen=sum(n_ongewogen)) %>%
  #Verwijder te lage aantallen
  mutate(across(.cols = "n_ongewogen" ,
                .fns = Vectorize(verwijder_kleine_aantallen), ongewogen = n_ongewogen)) %>%
  ungroup()

# aggregate(kubus_df$gmj_alcohollft_gem, list(kubus_df$gmj_opleidingsniveau), FUN=mean) 
# Vang nullen als er geen observaties zijn
kubus_df_tot[3][kubus_df_tot[3] == 0] <- missing_voor_privacy
kubus_df_tot[2][kubus_df_tot[3] == 0] <- missing_voor_privacy

# Voeg kolom toe met geoitem ggd
kubus_df_tot <- cbind(Geoitem = nr_regio, kubus_df_tot)
# Voeg kolom toe met geolevel ggd
kubus_df_tot <- cbind(Geolevel = "ggd", kubus_df_tot)
# Voeg kolom toe met jaarvariabele
kubus_df_tot <- cbind(Jaar = jaarvariabele, kubus_df_tot)

#Kolomnamen toewijzen 
names(kubus_df_tot) <- c("Jaar","Geolevel","Geoitem",crossings,glue("{variabele}_{variabel_labels}"),'n_ongewogen')

# Voeg de GGD getallen toe aan de gebiedscijfers
kubus_df <- rbind(kubus_df, kubus_df_tot)

n_labels <- length(variabel_labels)

#Excelbestand maken
workbook <- createWorkbook()

#Data toevoegen aan WB
addWorksheet(workbook, sheetName = "Data")
writeData(workbook,"Data",kubus_df)

#Definities data toevoegen aan WB
addWorksheet(workbook, sheetName = "Data_def")

writeData(workbook, "Data_def",
          
          cbind("col" = c(type_periode,"geolevel","geoitem",
                          #Alle crossings
                          crossings,
                          #Alle variabel-levels als kolommen
                          unlist(lapply(unname(variabel_labels), function(x){glue("{variabele}_{x}")})),
                          #Ongewogen kolom
                          glue("{variabele}_ONG")),
                "type" = c("period","geolevel","geoitem",
                           #Crossings zijn dimensies
                           rep("dim", length(crossings)),
                           #Variabel_levels (en de _ONG kolom) zijn variabelen
                           rep("var", n_labels + 1)))
          )

addWorksheet(workbook, sheetName = "Dimensies")
writeData(workbook,"Dimensies",
          
          cbind("Dimensiecode" = c(crossings),
                "Naam" = namen_crossings[k]))

#In De Swing viewer worden 'randtotalen' van crossings standaard als som uitgerekend.
#Een tabel over bijvoorbeeld de crossing leeftijd geeft dan de som van alle percentages
#als 'rijtotaal'. De som van de percentages voor verschillende leeftijden is een
#betekenisloos gegeven. Het gemiddelde percentage over alle leeftijden is wel nuttig.
#De sheet Dimension_levels dient om de default aggregatie over crossings aan te passen naar mean
addWorksheet(workbook, sheetName = "Dimension_levels")
writeData(workbook,"Dimension_levels",
          
          cbind("Dimlevel code" = c(crossings),
                "Name" = namen_crossings[k],
                "Dimension code" = c(crossings),
                "AggregateType" = "Mean"
          )
)

#Per crossing / dimensie een sheet toevoegen        
addWorksheet(workbook, sheetName = crossings)
writeData(workbook, crossings, 
          # Flexibele crossing labels invoegen
          cbind("Itemcode" = unlist(str_split(config_gem$labels_id[k], pattern = ",")),
                "Naam" = unlist(str_split(config_gem$labels_name[k], pattern = ",")),
                "Volgnr" = seq(1:length(unlist(str_split(config_gem$labels_name[k], pattern = ","))))
          )
)

#Tabblad Indicators; 
addWorksheet(workbook, sheetName = "Indicators")

writeData(workbook, "Indicators",
          
          cbind("Indicator code" = 
                  #Alle variabel-levels
                  c(
                    glue("{variabele}_{variabel_labels}"),
                    #ongewogen totaal
                    glue("{variabele}_ONG")
                  ),

                "Name" = 
                  #Alle variabel-levels
                  c(vars_labels[l],
                    "Totaal aantal ongewogen"
                  ), 
                
                #Percentage of personen
                #aantal rijen voor personen = levels variabele + 2(gewogen/ongewogen)
                "Unit" = c("Gemiddelde",
                  rep("personen",n_labels)
                        ),
              
                #Zelfde logica als "Unit"
                "Data type" = c("Mean",
                              rep("Numeric",n_labels)
                              ),
                                
                #Alleen het gemiddelde moet zichtbaar zijn.
                "Visible" = c(1,rep(0,n_labels)
                              ),

                #Treshold opgeven voor percentages. Als een selectie van
                #dimensies/crossings tot een groepsindeling leidt met minder observaties treshold 
                #Wordt dit afgeschermd in de mozaieken / viewer van Swing
                "Threshold value" = c(rep("",n_labels),min_observaties
                                      ),
                                      
                #Cube is overal 1
                "Cube" = c(rep(1,n_labels+1)
                           ),
                #RoundOff op  0,1
                "RoundOff" = c("0,1",
                               1
                ),
                #Source is overal hetzelfde
                "Source" = c(rep(bron,n_labels+1)
                             )
          )
)
                
saveWorkbook(workbook, file = glue("{basismap_output}/kubus_gemiddelde_{gebiedsniveau}_{variabele}_{crossings}.xlsx"), overwrite = TRUE)

  
#
##
## Platte data - randen (gemiddelde per gemeente)
## Totaal voor GGD regio uitgesplits en geheel.
#
#
# Selecteren vanuit DF naar kubus en gewogen getallen produceren.
kubus_df_plat <- data_totaal %>%
  filter(jaarvariabele %in% jaren_voor_analyse,
         !is.na(.[[variabele]]),
         !is.na(.[[gebiedsindeling]]),
         !is.na(weegfactor)) %>%
  mutate(n = 1,
         var = as.numeric(.[[variabele]],
                          levels = val_labels(.[[variabele]]),
                          labels = names(val_labels(.[[variabele]]))))%>%
  group_by(.[[jaarvariabele]],.[[gebiedsindeling]], var) %>%
  summarise(n_gewogen = sum(weegfactor),
            n_ongewogen = sum(n)) %>%
  ungroup()

# Bereken het gewogen gemiddelde per gebiedsniveau en per crossing, afronden op 1 cijfer achter de komma
kubus_df_plat <- kubus_df_plat %>%
  mutate(var_gewogen = var * n_gewogen) %>%
  mutate(var_ongewogen = var * n_ongewogen) %>%
  group_by(`.[[gebiedsindeling]]`) %>%
  summarise(Gew_ge=round(sum(var_gewogen)/sum(n_gewogen), digits = 1),
            n_ongewogen=sum(n_ongewogen)) %>%
  #Verwijder te lage aantallen
  mutate(across(.cols = "n_ongewogen" ,
                .fns = Vectorize(verwijder_kleine_aantallen), ongewogen = n_ongewogen)) %>%
  ungroup()

# Vang nullen als er geen observaties zijn
kubus_df_plat[3][kubus_df_plat[3] == 0] <- missing_voor_privacy
kubus_df_plat[2][kubus_df_plat[3] == 0] <- missing_voor_privacy

# Voeg kolom toe met geolevel op basis van gebiedsniveau
kubus_df_plat <- cbind(Geolevel = gebiedsniveau, kubus_df_plat)
# Voeg kolom toe met jaarvariabele
kubus_df_plat <- cbind(Jaar = jaarvariabele, kubus_df_plat)

#Kolomnamen toewijzen 
names(kubus_df_plat) <- c("Jaar","Geolevel","Geoitem",glue("{variabele}_{variabel_labels}"),'n_ongewogen')

## Toevoeging om het gemiddelde over de GGD regio te berekenen.
# Totaal voor GGD regio als geheel. Toevoeging 27 mei 2024; gebruik weegfactor_landelijk
kubus_df_plat_tot <- data_totaal %>%
  filter(jaarvariabele %in% jaren_voor_analyse,
         !is.na(.[[variabele]]),
         !is.na(.[[gebiedsindeling]]),
         !is.na(weegfactor_landelijk)) %>%
  mutate(n = 1,
         var = as.numeric(.[[variabele]],
                          levels = val_labels(.[[variabele]]),
                          labels = names(val_labels(.[[variabele]]))))%>%
  group_by(.[[jaarvariabele]], var) %>%
  summarise(n_gewogen = sum(weegfactor_landelijk),
            n_ongewogen = sum(n)) %>%
  ungroup()

# Bereken het gewogen gemiddelde voor ZL, afronden op 2 cijfers achter de komma
kubus_df_plat_tot <- kubus_df_plat_tot %>%
  mutate(var_gewogen = var * n_gewogen) %>%
  mutate(var_ongewogen = var * n_ongewogen) %>%
  summarise(Gew_ge=round(sum(var_gewogen)/sum(n_gewogen), digits = 2),
            n_ongewogen=sum(n_ongewogen)) %>%
  #Verwijder te lage aantallen
  mutate(across(.cols = "n_ongewogen" ,
                .fns = Vectorize(verwijder_kleine_aantallen), ongewogen = n_ongewogen)) %>%
  ungroup()

# aggregate(kubus_df$gmj_alcohollft_gem, list(kubus_df$gmj_opleidingsniveau), FUN=mean) 
# Vang nullen als er geen observaties zijn
kubus_df_plat_tot[2][kubus_df_plat_tot[2] == 0] <- missing_voor_privacy
kubus_df_plat_tot[1][kubus_df_plat_tot[2] == 0] <- missing_voor_privacy

# Voeg kolom toe met geoitem ggd
kubus_df_plat_tot <- cbind(Geoitem = nr_regio, kubus_df_plat_tot)
# Voeg kolom toe met geolevel ggd
kubus_df_plat_tot <- cbind(Geolevel = "ggd", kubus_df_plat_tot)
# Voeg kolom toe met jaarvariabele
kubus_df_plat_tot <- cbind(Jaar = jaarvariabele, kubus_df_plat_tot)

#Kolomnamen toewijzen 
names(kubus_df_plat_tot) <- c("Jaar","Geolevel","Geoitem",glue("{variabele}_{variabel_labels}"),'n_ongewogen')

# Voeg de GGD regio getallen toe aan de gebiedscijfers
kubus_df_plat <- rbind(kubus_df_plat, kubus_df_plat_tot)

#Excelbestand maken
workbook <- createWorkbook()

#Data toevoegen aan WB
addWorksheet(workbook, sheetName = "Data")
writeData(workbook,"Data",kubus_df_plat)

#Definities data toevoegen aan WB
addWorksheet(workbook, sheetName = "Data_def")

writeData(workbook, "Data_def",
          
          cbind("col" = c(type_periode,"geolevel","geoitem",
                          #Alle variabel-levels als kolommen
                          unlist(lapply(unname(variabel_labels), function(x){glue("{variabele}_{x}")})),
                          #Ongewogen kolom
                          glue("{variabele}_ONG")),
                "type" = c("period","geolevel","geoitem",
                           #Variabel_levels (en de _ONG kolom) zijn variabelen
                           rep("var", n_labels + 1)))
)


#Tabblad Indicators; 
addWorksheet(workbook, sheetName = "Indicators")

writeData(workbook, "Indicators",
          
          cbind("Indicator code" = 
                  #Alle variabel-levels
                  c(
                    glue("{variabele}_{variabel_labels}"),
                    #ongewogen totaal
                    glue("{variabele}_ONG")
                  ),
                
                "Name" = 
                  #Alle variabel-levels
                  c(vars_labels[l],
                    "Totaal aantal ongewogen"
                  ), 
                
                #Percentage of personen
                #aantal rijen voor personen = levels variabele + 2(gewogen/ongewogen)
                "Unit" = c("Gemiddelde",
                           rep("personen",n_labels)
                ),
                
                #Zelfde logica als "Unit"
                "Data type" = c("Mean",
                                rep("Numeric",n_labels)
                ),
                
                #Alleen het gemiddelde moet zichtbaar zijn.
                "Visible" = c(1,rep(0,n_labels)
                ),
                
                #Treshold opgeven voor percentages. Als een selectie van
                #dimensies/crossings tot een groepsindeling leidt met minder observaties treshold 
                #Wordt dit afgeschermd in de mozaieken / viewer van Swing
                "Threshold value" = c(rep("",n_labels),min_observaties
                ),
                
                #Cube is overal 1
                "Cube" = c(rep(1,n_labels+1)
                ),
                #RoundOff op  0,1
                "RoundOff" = c("0,1",
                               1
                ),
                #Source is overal hetzelfde
                "Source" = c(rep(bron,n_labels+1)
                )
          )
)

saveWorkbook(workbook, file = glue("{basismap_output}/kubus_gemiddelde_{gebiedsniveau}_{variabele}_plat.xlsx"), overwrite = TRUE)

  } else { # Niet gemeente, dus niet per gebied uitsplitsen, voor provincie of nederland
    # Selecteren vanuit DF naar kubus en gewogen getallen produceren.
    kubus_df <- data_totaal %>%
      filter(jaarvariabele %in% jaren_voor_analyse,
             !is.na(.[[variabele]]),
             !is.na(weegfactor),
             across(.cols = all_of(crossings) ,
                    .fns = ~ !is.na(.x))) %>%
      mutate(n = 1,
             var = as.numeric(.[[variabele]],
                              levels = val_labels(.[[variabele]]),
                              labels = names(val_labels(.[[variabele]]))))%>%
      group_by(.[[jaarvariabele]],across(all_of(crossings)), var) %>%
      summarise(n_gewogen = sum(weegfactor),
                n_ongewogen = sum(n)) %>%
      ungroup()
    
    #Functie om te kleine aantallen per groep te verwijderen --> Hier vallen de cellen weg!!!
    verwijder_kleine_aantallen <- function(x, ongewogen){if(ongewogen < minimum_obs_per_rij & ongewogen != missing_voor_privacy){missing_voor_privacy}else{x}}
    
    # Bereken het gewogen gemiddelde per gebiedsniveau en per crossing, afronden op 1 cijfer achter de komma
    kubus_df <- kubus_df %>%
      mutate(var_gewogen = var * n_gewogen) %>%
      mutate(var_ongewogen = var * n_ongewogen) %>%
      group_by(across(all_of(crossings))) %>%
      summarise(Gew_ge=round(sum(var_gewogen)/sum(n_gewogen), digits = 1),
                n_ongewogen=sum(n_ongewogen)) %>%
      #Verwijder te lage aantallen
      mutate(across(.cols = "n_ongewogen" ,
                    .fns = Vectorize(verwijder_kleine_aantallen), ongewogen = n_ongewogen)) %>%
      ungroup()
    
    # Vang nullen als er geen observaties zijn
    kubus_df[3][kubus_df[3] == 0] <- missing_voor_privacy
    kubus_df[2][kubus_df[3] == 0] <- missing_voor_privacy
    
    # Inbouwen vangnet minimum observaties per cel. Expliciet gemaakt!
    kubus_df <- kubus_df %>% mutate(kubus_df, n_ongewogen = case_when(n_ongewogen < minimum_obs_per_rij ~ missing_voor_privacy, TRUE ~ n_ongewogen))
    kubus_df <- kubus_df %>% mutate(kubus_df, Gew_ge = case_when(n_ongewogen < minimum_obs_per_rij ~ missing_voor_privacy, TRUE ~ Gew_ge))
    
    # Voeg kolom toe met geolevel ggd
    kubus_df <- cbind(Geoitem = nr_regio, kubus_df)
    # Voeg kolom toe met geolevel op basis van gebiedsniveau
    kubus_df <- cbind(Geolevel = gebiedsniveau, kubus_df)
    # Voeg kolom toe met jaarvariabele
    kubus_df <- cbind(Jaar = jaarvariabele, kubus_df)
    
    #Kolomnamen toewijzen 
    names(kubus_df) <- c("Jaar","Geolevel","Geoitem",crossings,glue("{variabele}_{variabel_labels}"),'n_ongewogen')
    
    n_labels <- length(variabel_labels)
    
    #Excelbestand maken
    workbook <- createWorkbook()
    
    #Data toevoegen aan WB
    addWorksheet(workbook, sheetName = "Data")
    writeData(workbook,"Data",kubus_df)
    
    #Definities data toevoegen aan WB
    addWorksheet(workbook, sheetName = "Data_def")
    
    writeData(workbook, "Data_def",
              
              cbind("col" = c(type_periode,"geolevel","geoitem",
                              #Alle crossings
                              crossings,
                              #Alle variabel-levels als kolommen
                              unlist(lapply(unname(variabel_labels), function(x){glue("{variabele}_{x}")})),
                              #Ongewogen kolom
                              glue("{variabele}_ONG")),
                    "type" = c("period","geolevel","geoitem",
                               #Crossings zijn dimensies
                               rep("dim", length(crossings)),
                               #Variabel_levels (en de _ONG kolom) zijn variabelen
                               rep("var", n_labels + 1)))
    )
    
    addWorksheet(workbook, sheetName = "Dimensies")
    writeData(workbook,"Dimensies",
              
              cbind("Dimensiecode" = c(crossings),
                    "Naam" = namen_crossings[k]))
    
    #In De Swing viewer worden 'randtotalen' van crossings standaard als som uitgerekend.
    #Een tabel over bijvoorbeeld de crossing leeftijd geeft dan de som van alle percentages
    #als 'rijtotaal'. De som van de percentages voor verschillende leeftijden is een
    #betekenisloos gegeven. Het gemiddelde percentage over alle leeftijden is wel nuttig.
    #De sheet Dimension_levels dient om de default aggregatie over crossings aan te passen naar mean
    addWorksheet(workbook, sheetName = "Dimension_levels")
    writeData(workbook,"Dimension_levels",
              
              cbind("Dimlevel code" = c(crossings),
                    "Name" = namen_crossings[k],
                    "Dimension code" = c(crossings),
                    "AggregateType" = "Mean"
              )
    )
    
    #Per crossing / dimensie een sheet toevoegen        
    addWorksheet(workbook, sheetName = crossings)
    writeData(workbook, crossings, 
              # Flexibele crossing labels invoegen
              cbind("Itemcode" = unlist(str_split(config_gem$labels_id[k], pattern = ",")),
                    "Naam" = unlist(str_split(config_gem$labels_name[k], pattern = ",")),
                    "Volgnr" = seq(1:length(unlist(str_split(config_gem$labels_name[k], pattern = ","))))
              )
    )
    
    #Tabblad Indicators; 
    addWorksheet(workbook, sheetName = "Indicators")
    
    writeData(workbook, "Indicators",
              
              cbind("Indicator code" = 
                      #Alle variabel-levels
                      c(
                        glue("{variabele}_{variabel_labels}"),
                        #ongewogen totaal
                        glue("{variabele}_ONG")
                      ),
                    
                    "Name" = 
                      #Alle variabel-levels
                      c(vars_labels[l],
                        "Totaal aantal ongewogen"
                      ), 
                    
                    #Percentage of personen
                    #aantal rijen voor personen = levels variabele + 2(gewogen/ongewogen)
                    "Unit" = c("Gemiddelde",
                               rep("personen",n_labels)
                    ),
                    
                    #Zelfde logica als "Unit"
                    "Data type" = c("Mean",
                                    rep("Numeric",n_labels)
                    ),
                    
                    #Alleen het gemiddelde moet zichtbaar zijn.
                    "Visible" = c(1,rep(0,n_labels)
                    ),
                    
                    #Treshold opgeven voor percentages. Als een selectie van
                    #dimensies/crossings tot een groepsindeling leidt met minder observaties treshold 
                    #Wordt dit afgeschermd in de mozaieken / viewer van Swing
                    "Threshold value" = c(rep("",n_labels),min_observaties
                    ),
                    
                    #Cube is overal 1
                    "Cube" = c(rep(1,n_labels+1)
                    ),
                    #RoundOff op  0,1
                    "RoundOff" = c("0,1",
                                   1
                    ),
                    #Source is overal hetzelfde
                    "Source" = c(rep(bron,n_labels+1)
                    )
              )
    )
    
    saveWorkbook(workbook, file = glue("{basismap_output}/kubus_gemiddelde_{gebiedsniveau}_{variabele}_{crossings}.xlsx"), overwrite = TRUE)
    

#
##
## Platte data - randen (gemiddelde per gebiedsindeling)
## Toevoeging om het gemiddelde over de GGD regio te berekenen.
## Totaal voor GGD regio uitgesplitst en geheel.
# Toevoeging 25 mei 2024; weegfactor_landelijk
#
# Selecteren vanuit DF naar kubus en gewogen getallen produceren.
kubus_df_plat <- data_totaal %>%
  filter(jaarvariabele %in% jaren_voor_analyse,
         !is.na(.[[variabele]]),
         !is.na(weegfactor_landelijk)) %>%
  mutate(n = 1,
         var = as.numeric(.[[variabele]],
                          levels = val_labels(.[[variabele]]),
                          labels = names(val_labels(.[[variabele]]))))%>%
  group_by(.[[jaarvariabele]], var) %>%
  summarise(n_gewogen = sum(weegfactor_landelijk),
            n_ongewogen = sum(n)) %>%
  ungroup()

# Bereken het gewogen gemiddelde per gebiedsniveau en per crossing, afronden op 1 cijfer achter de komma
kubus_df_plat <- kubus_df_plat %>%
  mutate(var_gewogen = var * n_gewogen) %>%
  mutate(var_ongewogen = var * n_ongewogen) %>%
  summarise(Gew_ge=round(sum(var_gewogen)/sum(n_gewogen), digits = 1),
            n_ongewogen=sum(n_ongewogen)) %>%
  #Verwijder te lage aantallen
  mutate(across(.cols = "n_ongewogen" ,
                .fns = Vectorize(verwijder_kleine_aantallen), ongewogen = n_ongewogen)) %>%
  ungroup()

# Vang nullen als er geen observaties zijn
kubus_df_plat[2][kubus_df_plat[2] == 0] <- missing_voor_privacy
kubus_df_plat[1][kubus_df_plat[2] == 0] <- missing_voor_privacy

# Voeg kolom toe met geolevel op basis van gebiedsniveau
kubus_df_plat <- cbind(Geoitem = nr_regio, kubus_df_plat)
# Voeg kolom toe met geolevel op basis van gebiedsniveau
kubus_df_plat <- cbind(Geolevel = gebiedsniveau, kubus_df_plat)
# Voeg kolom toe met jaarvariabele
kubus_df_plat <- cbind(Jaar = jaarvariabele, kubus_df_plat)

#Kolomnamen toewijzen 
names(kubus_df_plat) <- c("Jaar","Geolevel","Geoitem",glue("{variabele}_{variabel_labels}"),'n_ongewogen')


#Excelbestand maken
workbook <- createWorkbook()

#Data toevoegen aan WB
addWorksheet(workbook, sheetName = "Data")
writeData(workbook,"Data",kubus_df_plat)

#Definities data toevoegen aan WB
addWorksheet(workbook, sheetName = "Data_def")

writeData(workbook, "Data_def",
          
          cbind("col" = c(type_periode,"geolevel","geoitem",
                          #Alle variabel-levels als kolommen
                          unlist(lapply(unname(variabel_labels), function(x){glue("{variabele}_{x}")})),
                          #Ongewogen kolom
                          glue("{variabele}_ONG")),
                "type" = c("period","geolevel","geoitem",
                           #Variabel_levels (en de _ONG kolom) zijn variabelen
                           rep("var", n_labels + 1)))
)


#Tabblad Indicators; 
addWorksheet(workbook, sheetName = "Indicators")

writeData(workbook, "Indicators",
          
          cbind("Indicator code" = 
                  #Alle variabel-levels
                  c(
                    glue("{variabele}_{variabel_labels}"),
                    #ongewogen totaal
                    glue("{variabele}_ONG")
                  ),
                
                "Name" = 
                  #Alle variabel-levels
                  c(vars_labels[l],
                    "Totaal aantal ongewogen"
                  ), 
                
                #Percentage of personen
                #aantal rijen voor personen = levels variabele + 2(gewogen/ongewogen)
                "Unit" = c("Gemiddelde",
                           rep("personen",n_labels)
                ),
                
                #Zelfde logica als "Unit"
                "Data type" = c("Mean",
                                rep("Numeric",n_labels)
                ),
                
                #Alleen het gemiddelde moet zichtbaar zijn.
                "Visible" = c(1,rep(0,n_labels)
                ),
                
                #Treshold opgeven voor percentages. Als een selectie van
                #dimensies/crossings tot een groepsindeling leidt met minder observaties treshold 
                #Wordt dit afgeschermd in de mozaieken / viewer van Swing
                "Threshold value" = c(rep("",n_labels),min_observaties
                ),
                
                #Cube is overal 1
                "Cube" = c(rep(1,n_labels+1)
                ),
                #RoundOff op  0,1
                "RoundOff" = c("0,1",
                               1
                ),
                #Source is overal hetzelfde
                "Source" = c(rep(bron,n_labels+1)
                )
          )
)

saveWorkbook(workbook, file = glue("{basismap_output}/kubus_gemiddelde_{gebiedsniveau}_{variabele}_plat.xlsx"), overwrite = TRUE)   
}
}
}

