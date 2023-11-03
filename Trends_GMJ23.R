#
# Script voor het maken van trends over de jaren 2015 - 2019 - 2021 - 2023
#
# Werkgroep technische rapportage GezondheidsMonitor Jeugd 2023
# 
# Benodigdheden
library(dplyr)
library(readxl) 
library(haven)

setwd("C:/Documents/Profielen/GMJ2023")
# Vaststellen welke trends gemaakt moeten worden
# Indicatorenoverzicht trends laden en regel aanmaken per indicator in de kolom 'niveau'
trends <- read_excel('Indicatoren overzicht_GMJ_2023.xlsx', sheet = 'indicatoren trends') %>%
  mutate(niveau = strsplit(niveau, ", ")) %>%
  unnest(niveau)

# Trends maken
# Data 2023 laden
data2023 <- read_spss('SPSS_bestand_2023.sav')

# Data 2021 laden
data2021 <- read_spss('SPSS_bestand_2021.sav')

# Data 2019 laden
data2019 <- read_spss('SPSS_bestand_2019.sav')

# Data 2015 laden
data2015 <- read_spss('SPSS_bestand_2015.sav')

# Indien er trendbestand is met data voor alle jaren; data uit totaalbestand filteren.
# Pas de indicator jaar_variabele aan naar de kolom die het jaartal bevat.
# data <- read_spss('SPSS_trendsbestand_2023_2015.sav')
# data2023 <- data %>% filter(jaar_variabele == 2023)
# data2021 <- data %>% filter(jaar_variabele == 2021)
# data2019 <- data %>% filter(jaar_variabele == 2019)
# data2015 <- data %>% filter(jaar_variabele == 2015)


# Trendcijfers 2023
trends2023 <- trends %>%
  filter(!is.na(trends$weegfactor2023)) %>%
  select(indicator, niveau, weegfactor = weegfactor2023) %>%
  pmap(compute_mean, data = data, uitsplitsing = NA) %>% # compute_mean functie toepassen op het input object
  bind_rows() %>%
  mutate(name = str_replace(name, '_totaal', '')) %>%
  pivot_wider(names_from = name, values_from = value, values_fn = function(x) first(na.omit(x))) %>% # draaien van de output naar wide format
  mutate_at(vars(-1), round, 6) %>%
  setNames(c('niveau', paste0(names(.)[-1], '_2023')))

# Trendcijfers 2021
# Aanmaken van niveauvariabelen
data2021 <-  data2021 %>%
  mutate(totaal = 'totaal', # variabele aanmaken om het totaalgemiddelde te kunnen berekenen
         regio = ifelse(MIREB201 == regiocode, regionaam, NA), # variabele voor de regio aanmaken op basis van de eerder opgegeven regiocode en regionaam
         nederland = 'Nederland',
         Gemeentecode = ifelse(MIREB201 == regiocode, to_character(Gemeentecode), NA))

# Variabelen dichotomiseren die in het indicatorenoverzicht met een '_' en een getal in de indicatornaam 
if(any(trends$dichotomiseren == 1 & !is.na(trends$weegfactor2021))){
  data2021 <- dummy_cols(data2021,
                         select_columns = unique(str_extract(trends$indicator[trends$dichotomiseren == 1 & !is.na(trends$weegfactor2021)], '.+(?=_[0-9]+$)')),
                         ignore_na = T)
}


# Hercoderen van variabele met 8 = 'nvt' naar 0 zodat de percentages een weergave zijn van de totale groep
# Dit stukje code geeft een warning die kan worden genegeerd.
data2021 <- data2021 %>%
  mutate_at(c(data2021 %>%
                select(unique(trends$indicator[!is.na(trends$weegfactor2021)])) %>%
                val_labels() %>%
                str_detect('[Nn][\\.]?[Vv][\\.]?[Tt]') %>%
                unique(trends$indicator[!is.na(trends$weegfactor2021)])[.]), 
            list(~recode(., `8`= 0)))

# Berekenen trendcijfers
trends2021 <- trends %>%
  filter(!is.na(trends$weegfactor2021)) %>%
  select(indicator, niveau, weegfactor = weegfactor2021) %>%
  pmap(compute_mean, data = data2021, uitsplitsing = NA) %>% # compute_mean functie toepassen op het input object
  bind_rows() %>%
  mutate(name = str_replace(name, '_totaal', '')) %>%
  pivot_wider(names_from = name, values_from = value, values_fn = function(x) first(na.omit(x))) %>% # draaien van de output naar wide format
  mutate_at(vars(-1), round, 6) %>%
  setNames(c('niveau', paste0(names(.)[-1], '_2021')))

#
# Trendcijfers 2019
#
# Aanmaken van niveauvariabelen
data2019 <-  data2019 %>%
  mutate(totaal = 'totaal', # variabele aanmaken om het totaalgemiddelde te kunnen berekenen
         regio = ifelse(MIREB201 == regiocode, regionaam, NA), # variabele voor de regio aanmaken op basis van de eerder opgegeven regiocode en regionaam
         nederland = 'Nederland',
         Gemeentecode = ifelse(MIREB201 == regiocode, to_character(Gemeentecode), NA))

# Variabelen dichotomiseren die in het indicatorenoverzicht met een '_' en een getal in de indicatornaam 
if(any(trends$dichotomiseren == 1 & !is.na(trends$weegfactor2019))){
  data2019 <- dummy_cols(data2019,
                         select_columns = unique(str_extract(trends$indicator[trends$dichotomiseren == 1 & !is.na(trends$weegfactor2019)], '.+(?=_[0-9]+$)')),
                         ignore_na = T)
}


# Hercoderen van variabele met 8 = 'nvt' naar 0 zodat de percentages een weergave zijn van de totale groep
# Dit stukje code geeft een warning die kan worden genegeerd.
data2019 <- data2019 %>%
  mutate_at(c(data2019 %>%
                select(unique(trends$indicator[!is.na(trends$weegfactor2019)])) %>%
                val_labels() %>%
                str_detect('[Nn][\\.]?[Vv][\\.]?[Tt]') %>%
                unique(trends$indicator[!is.na(trends$weegfactor2021)])[.]), 
            list(~recode(., `8`= 0)))

# Berekenen trendcijfers
trends2019 <- trends %>%
  filter(!is.na(trends$weegfactor2019)) %>%
  select(indicator, niveau, weegfactor = weegfactor2019) %>%
  pmap(compute_mean, data = data2019, uitsplitsing = NA) %>% # compute_mean functie toepassen op het input object
  bind_rows() %>%
  mutate(name = str_replace(name, '_totaal', '')) %>%
  pivot_wider(names_from = name, values_from = value, values_fn = function(x) first(na.omit(x))) %>% # draaien van de output naar wide format
  mutate_at(vars(-1), round, 6) %>%
  setNames(c('niveau', paste0(names(.)[-1], '_2019')))

#
# Trendcijfers 2015
#
# Aanmaken van niveauvariabelen
data2015 <-  data2015 %>%
  mutate(totaal = 'totaal', # variabele aanmaken om het totaalgemiddelde te kunnen berekenen
         regio = ifelse(MIREB201 == regiocode, regionaam, NA), # variabele voor de regio aanmaken op basis van de eerder opgegeven regiocode en regionaam
         nederland = 'Nederland',
         Gemeentecode = ifelse(MIREB201 == regiocode, to_character(Gemeentecode), NA))

# Variabelen dichotomiseren die in het indicatorenoverzicht met een '_' en een getal in de indicatornaam 
if(any(trends$dichotomiseren == 1 & !is.na(trends$weegfactor2015))){
  data2015 <- dummy_cols(data2015,
                         select_columns = unique(str_extract(trends$indicator[trends$dichotomiseren == 1 & !is.na(trends$weegfactor2015)], '.+(?=_[0-9]+$)')),
                         ignore_na = T)
}


# Hercoderen van variabele met 8 = 'nvt' naar 0 zodat de percentages een weergave zijn van de totale groep
# Dit stukje code geeft een warning die kan worden genegeerd.
data2015 <- data2015 %>%
  mutate_at(c(data2015 %>%
                select(unique(trends$indicator[!is.na(trends$weegfactor2015)])) %>%
                val_labels() %>%
                str_detect('[Nn][\\.]?[Vv][\\.]?[Tt]') %>%
                unique(trends$indicator[!is.na(trends$weegfactor2021)])[.]), 
            list(~recode(., `8`= 0)))

# Berekenen trendcijfers
trends2015 <- trends %>%
  filter(!is.na(trends$weegfactor2015)) %>%
  select(indicator, niveau, weegfactor = weegfactor2015) %>%
  pmap(compute_mean, data = data2015, uitsplitsing = NA) %>% # compute_mean functie toepassen op het input object
  bind_rows() %>%
  mutate(name = str_replace(name, '_totaal', '')) %>%
  pivot_wider(names_from = name, values_from = value, values_fn = function(x) first(na.omit(x))) %>% # draaien van de output naar wide format
  mutate_at(vars(-1), round, 6) %>%
  setNames(c('niveau', paste0(names(.)[-1], '_2015')))

# Combinere alle trends en sorteren kolommen op naam. Gemiddelden afronden op 6 digits
trends_totaal <- trends2023 %>%
  left_join(trends2021, by = 'niveau') %>%
  left_join(trends2019, by = 'niveau')%>% 
  left_join(trends2015, by = 'niveau')%>% 
  select(niveau, sort(colnames(.))) %>% 
  mutate_at(vars(-niveau), round, 6) 


