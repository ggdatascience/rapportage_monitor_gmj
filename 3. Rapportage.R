# Script Rapportage

# 0. Voorbereiding --------------------------------------------------------

# Environment leegmaken
rm(list = ls())

# Libraries laden
# Deze libraries moeten eenmalig worden geinstalleerd met install.packages()
library(extrafont) # Package om het lettertype Century Gothic te gebruiken
library(haven) # Package Voor het laden van spss data bestanden
library(labelled) # Package om te werken met datalabels 
library(readxl) # Package om te werken met exceldocumenten
library(tidyverse) # Meer informatie over het tidyverse is te vinden op: https://www.tidyverse.org/
library(mschart) # Package voor het aanmaken van office grafieken
library(officer) # Package om Powerpoint (en andere MS Office) documenten aan te maken of te bewerken vanuit R

# 1. Dataconfiguratie ----------------------------------------------------

# Inladen van een trendbestand. Laad hier je eigen bestand in.
data <- read_spss('Testdata GMJ2023.sav')


# 2. Excelbestand met configuratie laden ----------------------------------

padnaam_configuratie <- '1. Configuratie.xlsx'

configuratie <- padnaam_configuratie %>% 
  excel_sheets() %>% 
  set_names() %>% 
  map(read_excel, path = padnaam_configuratie)


# 3. Grafiekinstellingen en opmaak ----------------------------------------

# Kleurpalet aanmaken
kleuren <- c('#009898', '#91caca', '#b9b441', '#dcd9a6', '#e8525f', '#e9acb0')

# Lettertype importeren. In de console wordt gevraagd of je dit wil doen run 'y' in de console om te importeren.
extrafont::font_import(pattern = 'GOTHIC')

# Lettertype bepalen
lettertype <- "Century Gothic"

# Thema voor staafgrafiek aanmaken
chart_theme <- mschart_theme(
  main_title = fp_text(font.size = 9, bold = TRUE, font.family = lettertype, color = 'black'),
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
grafiekstijl <- function(x, grafiektitel = NULL, datalabels = T, labelpositie = 'outEnd', labelkleur = 'black', ylimiet = 1) {
  
  x %>% chart_data_stroke('transparent') %>%
    chart_labels(title = grafiektitel, xlab = NULL, ylab = NULL) %>%
    chart_ax_x(major_tick_mark = 'none') %>%
    chart_ax_y(limit_min = 0, num_fmt = '0%', major_tick_mark = 'none', limit_max = ylimiet) %>%
    chart_data_labels(show_val = datalabels, num_fmt = '0%', position = labelpositie) %>%
    chart_labels_text(fp_text(font.size = 8, font.family = lettertype, color = labelkleur)) 
  
}


# 4. Functie maken om cijfers te berekenen --------------------------------

bereken_cijfers <- function(data,
                            basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                            omschrijving = NA, indicator, waarden = 1, valuelabel = NA, uitsplitsing = NA, groepering = NA, niveau, jaar = NA, 
                            var_jaar = 'Onderzoeksjaar', Nvar = 30, Ncel = 5, toetsen = F) {
  
  niveau_var <- {if(niveau == 'basis') basis else if(niveau == 'referentie') referentie}
  niveau_label <- {if(niveau == 'basis') basis_label else if(niveau == 'referentie') referentie_label}
  
  data %>%
    filter(.data[[niveau_var]] == niveau_label & .data[[var_jaar]] %in% jaar) %>%
    select(all_of(setdiff(c(var_jaar, indicator, uitsplitsing, groepering), NA))) %>%
    group_by(across(all_of(setdiff(c(var_jaar, indicator, uitsplitsing, groepering), NA)))) %>%
    tally() %>%
    drop_na() %>%
    {if(is.na(uitsplitsing) & is.na(groepering)) .
      else group_by(., across(all_of(setdiff(c(uitsplitsing, groepering), NA))))} %>%
    mutate(ntot = sum(n),
           nmin = ntot-n,
           val = ifelse(ntot < Nvar | n < Ncel | nmin < Ncel | n == ntot, NA, n/ntot) %>% as.numeric(),
           indicator = indicator, 
           omschrijving = omschrijving,
           niveau = {if(niveau == 'basis') 'School' else if(niveau == 'referentie') 'Regio'}) %>%
    filter_at(vars(all_of(indicator)), function(x) x %in% waarden) %>%
    ungroup() %>%
    rename(jaar = 1, aslabel = 2) %>%
    {if(!is.na(valuelabel)) mutate(., aslabel = valuelabel) else .} %>%
    {if(!is.na(uitsplitsing)) rename(., uitsplitsing = .data[[uitsplitsing]]) else .} %>%
    {if(!is.na(groepering)) rename(., groepering = .data[[groepering]]) else .} %>%
    mutate(across(!n & !ntot & !nmin & !val & !jaar, to_character)) %>%
    select(indicator, omschrijving, jaar, niveau, aslabel, everything())
}

bereken_cijfers(data = data,
                basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                omschrijving = NA, indicator, waarden = 1, valuelabel = NA, uitsplitsing = NA, groepering = NA, niveau, jaar = NA, 
                var_jaar = 'Onderzoeksjaar', Nvar = 30, Ncel = 5, toetsen = F)


# 5. Functies maken om content te genereren -------------------------------

# functie voor het berekenen van percentages
type_percentage <- function(data, 
                            basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                            indicator, waarden, niveau, jaar,
                            format = 'percentage') {
  data %>%
    bereken_cijfers(basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                    indicator = indicator, waarden = waarden, niveau = niveau, jaar = jaar) %>%
    mutate(., val = round(val*100)) %>%
    select(val) %>%
    unlist() %>%
    unname() %>%
    {if(is.na(.)) '-' 
      else if(format == 'percentage') paste0(., '%') 
      else if(format == 'getal') .}
}

# functie voor staafgrafiek
type_staafgrafiek <- function(data,
                              basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                              omschrijving = NA, indicator, waarden = 1, uitsplitsing = NA, groepering = NA, niveau, jaar) {
  
  if(niveau == 'basis' | niveau == 'referentie') {
    
  data %>%
  bereken_cijfers(basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                  indicator = indicator, waarden = waarden, niveau = niveau, jaar = jaar, omschrijving = omschrijving, uitsplitsing = uitsplitsing, groepering = groepering) %>%
    {if(!is.na(uitsplitsing) & !is.na(groepering))
          ms_barchart(., x = 'uitsplitsing', y = 'val', group = 'groepering') %>%
          chart_data_fill(kleuren[1:(val_labels(data[[groepering]]) %>% length())] %>% set_names(val_labels(data[[groepering]]) %>% names())) %>%
          set_theme(chart_theme) %>%
          grafiekstijl(grafiektitel = omschrijving)
  
      else if(!is.na(uitsplitsing) & is.na(groepering))
          ms_barchart(., x = 'uitsplitsing', y = 'val') %>%
          chart_data_fill(kleuren[1]) %>%
          set_theme(chart_theme) %>%
          chart_theme(legend_position = 'n') %>%
        grafiekstijl(grafiektitel = omschrijving)
  
      else if(is.na(uitsplitsing) & !is.na(groepering))
          ms_barchart(., x = 'omschrijving', y = 'val', group = 'groepering') %>%
          chart_data_fill(kleuren[1:(val_labels(data[[groepering]]) %>% length())] %>% set_names(val_labels(data[[groepering]]) %>% names())) %>%
          set_theme(chart_theme) %>%
          chart_ax_x(tick_label_pos = 'none') %>%
        grafiekstijl(grafiektitel = omschrijving)
  
      else
        ms_barchart(., x = 'omschrijving', y = 'val') %>%
          chart_data_fill(kleuren[1]) %>%
          set_theme(chart_theme) %>%
          chart_theme(legend_position = 'n') %>%
          chart_ax_x(tick_label_pos = 'none') %>%
          grafiekstijl(grafiektitel = omschrijving)
    }
    
  } else if(niveau == 'basis en referentie') {
    
    data.frame(niveau = c('basis', 'referentie'),
               basis = c(basis, NA),
               basis_label = c(basis_label, NA),
               referentie = c(NA, referentie),
               referentie_label = c(NA, referentie_label)) %>%
      pmap(bereken_cijfers, data = data, indicator = indicator, waarden = waarden, jaar = jaar, 
           omschrijving = omschrijving, uitsplitsing = uitsplitsing, groepering = groepering) %>%
      reduce(bind_rows) %>%
      mutate(niveau = factor(niveau, levels = c('School', 'Regio'), ordered = T)) %>%
      ms_barchart(x = 'omschrijving', y = 'val', group = 'niveau') %>%
      chart_data_fill(kleuren[c(1,5)] %>% set_names(c('School', 'Regio'))) %>%
      set_theme(chart_theme) %>%
      chart_theme(legend_position = 'b') %>%
      chart_ax_x(tick_label_pos = 'none') %>%
      chart_settings(overlap = -20) %>%
      grafiekstijl(grafiektitel = omschrijving)
    
  }
}

# functie voor trendgrafiek
type_trendgrafiek <- function(data,
                              basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                              omschrijving = NA, indicator, waarden = 1, niveau, jaar) {
  
  data.frame(niveau = c('basis', 'referentie'),
             basis = c(basis, NA),
             basis_label = c(basis_label, NA),
             referentie = c(NA, referentie),
             referentie_label = c(NA, referentie_label)) %>%
    pmap(bereken_cijfers, data = data, indicator = indicator, waarden = waarden, jaar = jaar, 
         omschrijving = omschrijving) %>%
    reduce(bind_rows) %>%
    mutate(niveau = factor(niveau, levels = c('School', 'Regio'), ordered = T),
           jaar = as.character(jaar)) %>%
    ms_linechart(x = 'jaar', y = 'val', group = 'niveau') %>%
    set_theme(chart_theme) %>%
    grafiekstijl(grafiektitel = omschrijving, datalabels = F, labelpositie = 't') %>%
    chart_data_fill(kleuren[c(1,5)] %>% set_names(c('School', 'Regio'))) %>%
    chart_data_stroke(kleuren[c(1,5)] %>% set_names(c('School', 'Regio'))) %>%
    chart_data_symbol('none')
}

# functie voor combigrafiek
type_combi <- function(data,
                       basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                       omschrijving = NA, indicator, waarden = 1, uitsplitsing = NA, groepering = NA, niveau, jaar,
                       valuelabel = NA, horizontaal = F) {
  
  data.frame(indicator = indicator,
             valuelabel = valuelabel) %>%
    pmap(bereken_cijfers, data = data, basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
      waarden = waarden, niveau = niveau, jaar = jaar, omschrijving = omschrijving, uitsplitsing = uitsplitsing, groepering = groepering) %>%
    reduce(bind_rows) %>%
    {if(!is.na(groepering)) 
      ms_barchart(.,x = 'aslabel', y = 'val', group = 'groepering') %>%
        chart_data_fill(kleuren[1:(val_labels(data[[groepering]]) %>% length())] %>% set_names(val_labels(data[[groepering]]) %>% names())) %>%
        set_theme(chart_theme)
        
      else ms_barchart(.,x = 'aslabel', y = 'val') %>%
        chart_data_fill(kleuren[1]) %>%
        set_theme(chart_theme) %>%
        chart_theme(legend_position = 'n')} %>%
    grafiekstijl(grafiektitel = omschrijving) %>%
    {if(horizontaal == T) chart_settings(., dir = "horizontal") else .}

}


# 6. Functies aanmaken om content te genereren en te plaatsen -------------

# Functie aanmaken die op basis van de kolom 'type' uit de configuratie content genereert.
content_plaatsen <- function(data, template, 
                             basis, basis_label, referentie, referentie_label, 
                             omschrijving, type, indicator, waarden, valuelabel, uitsplitsing, groepering, niveau, jaar, index, label) {
  
  if(type == 'rapportnaam') {
  
    value <- basis_label
      
  } else if(type == 'datum') {
    
    value <- Sys.Date() %>% format(format = "%d-%m-%Y") %>% as.character()
    
  } else if(type == 'percentage') {
    
    value <- type_percentage(data, 
                             basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                             indicator = indicator, waarden = waarden, niveau = niveau, jaar = jaar,
                             format = 'percentage')
  
  } else if(type == 'staafgrafiek') {
    
    value <- type_staafgrafiek(data = data,
                               basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                               omschrijving = omschrijving, indicator = indicator, uitsplitsing = uitsplitsing, groepering = groepering, 
                               waarden = waarden, niveau = niveau, jaar = jaar)
    
  } else if(type == 'trendgrafiek') {
    
    value <- type_trendgrafiek(data = data,
                               basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                               omschrijving = omschrijving, indicator = indicator, waarden = 1, niveau = niveau, jaar = jaar)
    
  } else if(type == 'combi') {
    
    value <- type_combi(data = data,
                        basis = basis, basis_label = basis_label, valuelabel = valuelabel,
                        omschrijving = omschrijving, indicator = indicator, uitsplitsing = uitsplitsing, groepering = groepering, waarden = waarden, niveau = niveau, jaar = jaar) 
  
  } else if(type == 'combi liggend') {
    
    value <- type_combi(data = data, horizontaal = T,
                        basis = basis, basis_label = basis_label, valuelabel = valuelabel,
                        omschrijving = omschrijving, indicator = indicator, uitsplitsing = uitsplitsing, groepering = groepering, waarden = waarden, niveau = niveau, jaar = jaar) 
    
  } else if(type == 'vergelijking') {
    
    value <- type_vergelijking(data = data, basis_label = basis_label, indicator = indicator, 
                      omschrijving = omschrijving, uitsplitsing = uitsplitsing)
  
  } else if(type == 'staafgrafiek gestapeld') {
    
    value <- type_staafgrafiek_gestapeld(data = data, basis_label = basis_label, indicator = indicator, 
                      omschrijving = omschrijving, uitsplitsing = uitsplitsing)
    
  } else if(type == 'top 3') {
    
    value <- type_top3()
    
  }
  
  on_slide(template, index = index) %>%
    ph_with(value = value, location = ph_location_label(ph_label = label))
  
}


# 7. Rapportage maken -----------------------------------------------------

# Functie aanmaken om de rapportage te maken. 
rapportage_maken <- function(data, template, configuratie, basis, basis_label, referentie, referentie_label, slideconfiguratie) {
  
  slideconfiguratie <- configuratie[[slideconfiguratie]] %>%
    select(-vraag) %>%
    mutate(across(c(indicator, valuelabel, niveau, jaar), ~ .x %>% str_split('; ')))
  
  template <- read_pptx(template)
  pwalk(.l = slideconfiguratie, .f = content_plaatsen, data = data, template = template, basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label)
  print(template, paste0('Rapport ', basis_label, '.pptx'))

}


# 8. Rapportages maken op basis van rapportconfiguratie -------------------

pmap(.l = configuratie$Rapportconfiguratie, 
     .f = rapportage_maken, data = data, configuratie = configuratie, template = '2. Template.pptx')
