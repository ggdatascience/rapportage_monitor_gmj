# Script Rapportage

# 0. Voorbereiding --------------------------------------------------------

# Environment leegmaken
rm(list = ls())

# Libraries laden
# Deze libraries moeten eenmalig worden geinstalleerd met install.packages()
library(haven) # Package Voor het laden van spss data bestanden
library(labelled) # Package om te werken met datalabels 
library(readxl) # Package om te werken met exceldocumenten
library(tidyverse) # Meer informatie over het tidyverse is te vinden op: https://www.tidyverse.org/
library(mschart) # Package voor het aanmaken van office grafieken
library(officer) # Package om Powerpoint (en andere MS Office) documenten aan te maken of te bewerken vanuit R


# 1. Dataconfiguratie ----------------------------------------------------

# Als je een trendbestand hebt met alle jaren 
padnamen <- c('PROEFBESTAND_GM Jeugd 2023_anoniem_met indicatoren.sav')

# Als je losse bestanden hebt per jaar
padnamen <- c('GMJ2023.sav',
              'GMJ2021.sav',
              'GMJ2019.sav',
              'GMJ2015.sav')

# Inladen van data
{ if(length(padnamen) == 1) {

  trendsdata <- haven::read_spss(padnamen)
  
  data <- trendsdata %>% filter(Onderzoeksjaar == 2023)
  
  }
  
  else { databestanden <- apply(padnamen, haven::read_spss)
  
  trendsdata <- lapply(databestanden, left_join, by = c('Onderzoeksjaar'))
  
  data <- trendsdata %>% filter(jaar == 2023)
  
  }
}


# databewerkingen uitvoeren
data <- data %>%
  mutate(schoolcode = labelled(sample(x = c(23001, 23002, 23003, 23004, 23005), size = nrow(data), replace = T),
                               c('School 1' = 23001, 'School 2' = 23002, 'School 3' = 23003, 'School 4' = 23004, 'School 5' = 23005))) %>%
  mutate(RAPPORT = to_character(schoolcode),
         KLAS = ifelse(KLAS == 0, NA, KLAS) %>% labelled(c('KLAS 2' = 1, 'KLAS 4' = 2)),
         GENDER = ifelse(GENDER == 9, NA, KLAS) %>% labelled(c('Jongen' = 1, 'Meisje' = 2)),
         MBOKK3S31 = ifelse(MBOKK3S31 == 8 | MBOKK3S31 == 9, NA, MBOKK3S31) %>% labelled(c('Vmbo' = 1, 'Havo/Vwo' = 2)))

# Testen van bovenstaande met bestanden van voorgaande jaren

# 2. Excelbestand met configuratie laden ----------------------------------
rapportconfiguratie <- readxl::read_xlsx("1. Configuratie.xlsx", sheet = 'Rapportconfiguratie')
slideconfiguratie <- readxl::read_xlsx("1. Configuratie.xlsx", sheet = 'Slideconfiguratie 1') %>%
  # slice(1:19) %>%
  # filter(werkt == 1) %>%
  select(-vraag, -werkt, -Opmerkingen_tijdelijk) %>%
  mutate(across(c(indicator, labels, niveau), ~ .x %>% str_split('; ')))

# 3. Grafiekinstellingen en opmaak ----------------------------------------

kleuren <- c('#e8525f', '#009898', 'purple', 'darkgreen', 'beige', 'orange')

lettertype <- "TT Arial"

# Thema voor staafgrafiek aanmaken
bartheme <- mschart_theme(
  main_title = fp_text(font.size = 12, bold = TRUE, font.family = lettertype, color = 'black'),
  axis_title_x = fp_text(font.size = 9, font.family = lettertype),
  axis_title_y = fp_text(font.size = 9, font.family = lettertype),
  axis_text_x = fp_text(font.size = 9, font.family = lettertype),
  axis_text_y = fp_text(font.size = 9, font.family = lettertype),
  legend_text = fp_text(font.size = 9, font.family = lettertype),
  grid_major_line = fp_border(style = "none"),
  axis_ticks_x =  fp_border(width = 1),
  axis_ticks_y =  fp_border(width = 1),
  legend_position = 'b'
)

# Grafiekinstellingen bepalen
grafiekstijl <- function(x, grafiektitel = NULL, labelpositie = 'outEnd', labelkleur = 'black', ylimiet = 1) {
  
  x %>% chart_data_stroke('transparent') %>%
    chart_labels(title = grafiektitel, xlab = NULL, ylab = NULL) %>%
    chart_ax_x(major_tick_mark = 'none') %>%
    chart_ax_y(limit_min = 0, num_fmt = '0%', major_tick_mark = 'none', limit_max = ylimiet) %>%
    chart_data_labels(show_val = T, num_fmt = '0%', position = labelpositie) %>%
    chart_labels_text(fp_text(font.size = 8, font.family = lettertype, color = labelkleur)) 
  
}


# 4. Functie maken om cijfers te berekenen --------------------------------

bereken_cijfers <- function(data, rapportnaam, indicator, waarden = 1, omschrijving = NA, uitsplitsing = NA, groepering = NA, Nvar = 30, Ncel = 5, toetsen = F) {
  
  result <- data %>%
    filter(RAPPORT == rapportnaam) %>%
    select(all_of(setdiff(c(indicator, uitsplitsing, groepering), NA))) %>%
    group_by(across(all_of(setdiff(c(indicator, uitsplitsing, groepering), NA)))) %>%
    tally() %>%
    drop_na() %>%
    {if(is.na(uitsplitsing) & is.na(groepering)) . 
      else group_by(., across(all_of(setdiff(c(uitsplitsing, groepering), NA))))} %>%
    mutate(ntot = sum(n),
           val = ifelse(ntot < Nvar | min(n) < Ncel | n == ntot, NA, n/ntot) %>% as.numeric()) %>%
    filter_at(vars(all_of(indicator)), function(x) x %in% waarden) %>%
    ungroup() %>%
    mutate(across(!n & !ntot & !val, to_character),
           n_min = ntot - n) %>%
    rename(var = 1) %>%
    mutate(var = omschrijving)

  # Perform chi-squared test
  
  if(toetsen == T){
    
  chi_squared_result <- chisq.test(result[, c("n", "n_min")])

  hoogste_waarde <- which.max(result$val)
  rij_met_hoogste_waarde <- result[hoogste_waarde, ]
  laagste_waarde <- which.min(result$val)
  rij_met_laagste_waarde <- result[laagste_waarde, ]

  # Extract 'omschrijving' for highest and lowest values
  hoogste_omschrijving <- rij_met_hoogste_waarde$var
  laagste_omschrijving <- rij_met_laagste_waarde$var
  hoogste_categorie <- rij_met_hoogste_waarde[2]
  laagste_categorie <- rij_met_laagste_waarde[2]

  # Add chi-squared test results to the result dataframe
  result$p_val <- chi_squared_result$p.value
  result$chi_val <- chi_squared_result$statistic
  result$df <- chi_squared_result$parameter

  }

  return(result)
  
}



# bereken_cijfers functie testen
bereken_cijfers(data, 'School 1', 'SBOSK3S4', waarden = c(1,2), omschrijving = 'Schoolbeleving', uitsplitsing = 'GENDER', groepering = 'KLAS', toetsen = F)
# bereken_cijfers(data, 'School 1', 'GELUK', omschrijving = 'Voelt zich gelukkig', toetsen = T)
# bereken_cijfers(data, 'School 1', 'GELUK', omschrijving = 'Voelt zich gelukkig', uitsplitsing = 'JAAR', toetsen = F)
# bereken_cijfers(data, 'School 1', 'GELUK', omschrijving = 'Voelt zich gelukkig', groepering = 'KLAS', toetsen = T)
# bereken_cijfers(data, 'School 1', 'GELUK', omschrijving = 'Voelt zich gelukkig', uitsplitsing = 'GESLACHT', groepering = 'KLAS')

# Voor meerdere splitsingen is de toegevoegde significantietoets niet geschikt. 
# bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig', uitsplitsing = 'GESLACHT', groepering = 'KLAS')
# bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig', uitsplitsing = 'GESLACHT', groepering = 'SCHOOL')


# 5. Functies maken om content te genereren -------------------------------

# functie voor losse cijfers
type_percentage <- function(data, rapportnaam, indicator, waarden, format = 'percentage', Nvar = 30, Ncel = 5) {
  
  data %>%
    bereken_cijfers(rapportnaam = rapportnaam, indicator = indicator, waarden = waarden, Nvar = Nvar, Ncel = Ncel) %>%
    mutate(., val = round(val*100)) %>%
    select(val) %>%
    unlist() %>%
    unname() %>%
    {if(is.na(.)) NA 
      else if(format == 'percentage') paste0(., '%') 
      else if(format == 'getal') .}
}

# type_percentage functie testen
# type_percentage(data, 'School 1', 'MBGSK3S2', waarden = 3)


# functie voor statistische toets tussen twee groepen
type_vergelijking <- function(data, rapportnaam, indicator, omschrijving, uitsplitsing, Nvar = 30, Ncel = 5, toetsen = T) {
  
  data %>%
    bereken_cijfers(rapportnaam = rapportnaam, indicator = indicator, omschrijving = omschrijving, uitsplitsing = uitsplitsing, Nvar = Nvar, Ncel = Ncel, toetsen = toetsen) %>%
    mutate(percentage = paste0(round(val*10, 0), '%')) %>%
    rename(uitsplitsing = 2) %>%
    select(-n, -ntot, -val, -n_min, -chi_val, -df) %>%
    pivot_wider(names_from = uitsplitsing,
                values_from = percentage) %>%
    mutate(var = str_replace(var, paste0('\\[', names(.)[3], '\\]'), unlist(.[1,3])), 
           var = str_replace(var, paste0('\\[', names(.)[4], '\\]'), unlist(.[1,4])),
           var = {if(p_val >= 0.05) str_replace(var, '\\[hoger dan/lager dan/vergelijkbaar met\\]', 'vergelijkbaar met')
             else if(p_val < 0.05) str_replace(var, '\\[hoger dan/lager dan/vergelijkbaar met\\]', ifelse(unlist(.[1,3]) > unlist(.[1,4]), 'hoger dan', 'lager dan'))
           }) %>%
    pull(var) 
    
}

# type_vergelijking functie testen
# type_vergelijking(data, 'School 1', 'GELUK', 
#                   omschrijving = 'Van de meisjes voelt [Meisje] zich gelukkig en dit percentage is [hoger dan/lager dan/vergelijkbaar met] jongens, waarvan [Jongen] zich gelukkig voelt.', 
#                   uitsplitsing = 'GESLACHT', 
#                   toetsen = T)


# functie voor staafgrafiek
type_staafgrafiek <- function(data, rapportnaam, indicator, omschrijving = NA, uitsplitsing = NA, groepering = NA, Nvar = 30, Ncel = 5) {
  
  data %>%
    bereken_cijfers(indicator = indicator, rapportnaam = rapportnaam, omschrijving = omschrijving, uitsplitsing = uitsplitsing, groepering = groepering, Nvar = Nvar, Ncel = Ncel) %>%
    {if(!is.na(uitsplitsing) & !is.na(groepering))
        rename(., uitsplitsing = 2, groepering = 3) %>%
          ms_barchart(x = 'uitsplitsing', y = 'val', group = 'groepering') %>%
          chart_data_fill(kleuren[1:(val_labels(data[[groepering]]) %>% length())] %>% set_names(val_labels(data[[groepering]]) %>% names())) %>%
          set_theme(bartheme)
      
      else if(!is.na(uitsplitsing) & is.na(groepering)) 
        rename(., uitsplitsing = 2) %>%
          ms_barchart(x = 'uitsplitsing', y = 'val') %>%
          chart_data_fill(kleuren[1]) %>%
          set_theme(bartheme) %>%
          chart_theme(legend_position = 'n')
      
      else if(is.na(uitsplitsing) & !is.na(groepering)) 
        rename(., groepering = 2) %>%
          ms_barchart(x = 'var', y = 'val', group = 'groepering') %>%
          chart_data_fill(kleuren[1:(val_labels(data[[groepering]]) %>% length())] %>% set_names(val_labels(data[[groepering]]) %>% names())) %>%
          set_theme(bartheme)
      
      else
        ms_barchart(., x = 'var', y = 'val') %>%
          chart_data_fill(kleuren[1]) %>%
          set_theme(bartheme) %>%
          chart_theme(legend_position = 'n') %>%
          chart_ax_x(tick_label_pos = 'none')
      } %>%
    grafiekstijl(grafiektitel = omschrijving)

}

# type_grafiek() functie testen
bereken_cijfers(data, 'School 1', 'SBOSK3S4', omschrijving = 'Schoolbeleving', uitsplitsing = 'GENDER', groepering = 'KLAS') %>%
  print(preview = T)

# type_staafgrafiek(data, rapportnaam = 'School 2', 'MBGSK3S2', omschrijving = 'Schoolbeleving', uitsplitsing = NA, groepering = NA) %>%
#   print(preview = T)
# type_staafgrafiek(data, rapportnaam = 'School 2', 'GELUK', omschrijving = 'Schoolbeleving', uitsplitsing = 'GESLACHT', groepering = NA) %>%
#   print(preview = T)
# type_staafgrafiek(data, rapportnaam = 'School 2', 'GELUK', omschrijving = 'Schoolbeleving', uitsplitsing = NA, groepering = 'OPLEIDING') %>%
#   print(preview = T)
# type_staafgrafiek(data, rapportnaam = 'School 2', 'GELUK', omschrijving = 'Schoolbeleving', uitsplitsing = 'GESLACHT', groepering = 'KLAS') %>%
#   print(preview = T)

# functie voor gestapelde staafgrafiek
# type_staafgrafiek_gestapeld <- function(data, rapportnaam, indicator, omschrijving = NA, uitsplitsing = NA, groepering = NA, Nvar = 30, Ncel = 5) {
#   
#   data %>%
#     bereken_cijfers(indicator = indicator, rapportnaam = rapportnaam, omschrijving = omschrijving, uitsplitsing = uitsplitsing, groepering = groepering, Nvar = Nvar, Ncel = Ncel)
#   
# }
# 
# type_staafgrafiek_gestapeld(data, rapportnaam = 'School 2', 'MBGSK3S2', omschrijving = 'Schoolbeleving', uitsplitsing = NA, groepering = NA) %>%
#   print(preview = T)

# functie voor combigrafiek
type_combi <- function(data, rapportnaam, indicator, omschrijving = NA, uitsplitsing = NA, groepering = NA, Nvar = 30, Ncel = 5, horizontaal = F) {
  
  data.frame(omschrijving = unlist(omschrijving),
             indicator = unlist(indicator)) %>%
    pmap(bereken_cijfers, data = data, rapportnaam = rapportnaam, uitsplitsing = uitsplitsing, groepering = groepering, Nvar = Nvar, Ncel = Ncel) %>%
    reduce(bind_rows) %>%
    rename(var = 1, uitsplitsing = 2) %>%
    ms_barchart(x = 'var', y = 'val', group = 'uitsplitsing') %>%
    chart_data_fill(kleuren[1:(val_labels(data[[uitsplitsing]]) %>% length())] %>% set_names(val_labels(data[[uitsplitsing]]) %>% names())) %>%
    set_theme(bartheme) %>%
    grafiekstijl() %>%
    {if(horizontaal == T) chart_settings(., dir = "horizontal") else .}

}

# type_combi functie testen
# type_combi(data = data, rapportnaam = 'School 1', indicator = slideconfiguratie$indicator[14], omschrijving = slideconfiguratie$omschrijving[14], uitsplitsing = 'KLAS') %>%
#   print(preview = T)


# 6. Functies aanmaken om content te genereren en te plaatsen -------------

# Functie aanmaken die op basis van de kolom 'type' uit de configuratie content genereert.
content_genereren <- function(data, rapportnaam, omschrijving, indicator, waarden, uitsplitsing, groepering, vergelijking, type) {
  
  if(type == 'rapportnaam') {
  
    rapportnaam
      
  } else if(type == 'percentage') {
    
    type_percentage(data = data, rapportnaam = rapportnaam, waarden = waarden, indicator = indicator)
  
  } else if(type == 'staafgrafiek') {
    
    type_staafgrafiek(data = data, rapportnaam = rapportnaam, indicator = indicator, 
                      omschrijving = omschrijving, uitsplitsing = uitsplitsing, groepering = groepering)
    
  } else if(type == 'vergelijking') {
    
    type_vergelijking(data = data, rapportnaam = rapportnaam, indicator = indicator, 
                      omschrijving = omschrijving, uitsplitsing = uitsplitsing)
  
  } else if(type == 'staafgrafiek gestapeld') {
    
    type_staafgrafiek_gestapeld(data = data, rapportnaam = rapportnaam, indicator = indicator, 
                      omschrijving = omschrijving, uitsplitsing = uitsplitsing)
    
  } else if(type == 'combi') {
    
    type_combi(data = data, rapportnaam = rapportnaam, indicator = indicator, omschrijving = omschrijving, 
               uitsplitsing = uitsplitsing, groepering = groepering)
  
  } else if(type == 'combi liggend') {
    
    type_combi(data = data, rapportnaam = rapportnaam, indicator = indicator, omschrijving = omschrijving, 
               uitsplitsing = uitsplitsing, groepering = groepering, horizontaal = T)
  }
}

# Testen van content_generen() functie
# content_genereren(data = data, rapportnaam = 'School 2', omschrijving = 'bla', indicator = 'GELUK', groepering = 'GESLACHT', type = 'percentage')

# Functie aanmaken om de content te plaatsen. Eerst wordt de content gegenereert met de 
# content_genereren() functie. Daarna wordt de content op basis van de configuratie
# op de juiste slide en de juiste aanduiding geplaatst.
content_plaatsen <- function(data, template, rapportnaam, omschrijving, indicator, waarden, uitsplitsing, groepering, vergelijking, type, index, label) {
  
  value <- content_genereren(data = data, rapportnaam = rapportnaam, omschrijving = omschrijving, indicator = indicator, waarden = waarden, uitsplitsing = uitsplitsing,
                             groepering = groepering, vergelijking = vergelijking, type = type)
  
  on_slide(template, index = index) %>%
    ph_with(value = value, location = ph_location_label(ph_label = label))
 
}


# 7. Rapportage maken -----------------------------------------------------

# Functie aanmaken om de rapportage te maken. 
rapportage_maken <- function(template, slideconfiguratie, data, rapportnaam) {
  
  template <- read_pptx(template)
  pwalk(.l = slideconfiguratie, .f = content_plaatsen, data = data, template = template, rapportnaam = rapportnaam)
  print(template, paste0('Rapport ', rapportnaam, '.pptx'))
  
}


# 8. Rapportage maken via functie -----------------------------------------

rapportage_maken(template = '2. Template.pptx',
                 slideconfiguratie = slideconfiguratie,
                 data = data,
                 rapportnaam = 'School 1')

# 9. Rapportages maken op basis van rapportconfiguratie -------------------

sapply(X = rapportconfiguratie$rapport, FUN = rapportage_maken, template = '2. Template.pptx',
       slideconfiguratie = slideconfiguratie, data = data)
