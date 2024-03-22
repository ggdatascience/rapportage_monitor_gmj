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
library(flextable) # Package om tabellen te maken

# Working directory bepalen
# Je kunt op meerdere manieren de werkmap bepalen waarin je documenten staan.
# In het paneel met het tabbled Files (meestal het paneel rechtsonder) zie je de map die R als werkmap beschouwd.
# Je kunt met de functie getwd() de padnaam zien van je werkmap.
getwd()

# De werkmap is meestal de map van waaruit je een R script opent. Is dit niet zo dan kun je met setwd() je werkmap bepalen.
setwd('C:/padnaam')
# De werkmap bepalen kan ook door bovenin te klikken op "Session"--> "Set Working Directory" --> "Choose Directory" (oftewel Ctrl + Shift +H)

# 1. Data inladen ---------------------------------------------------------

# Inladen van een trendbestand. Laad hier je eigen bestand in.
data <- read_spss('data.sav')


# 2. Excelbestand met configuratie laden ----------------------------------

# Padnaam opgeven van de Excel configuratie. Staat dit bestand in je working directory dan hoef je alleen de bestandsnaam
# op te geven (inclusief de bestandsextensie)
padnaam_configuratie <- '1. Configuratie - Gemeente.xlsx'


# 3. Grafiekinstellingen en opmaak ----------------------------------------

# Lettertype importeren. In de console wordt gevraagd of je dit wil doen run 'y' in de console om te importeren
extrafont::font_import(pattern = 'GOTHIC')

# Check om te zien of het lettertype goed is ingeladen (check of het lettertype Century Gothic ertussen staat)
windowsFonts()

# Grafiekinstellingen bepalen
grafiekstijl <- function(grafiek, 
                         lettertype = "Century Gothic", 
                         grafiektitel = NULL, 
                         legendapositie = 'b', 
                         datalabels = T, 
                         labelpositie = 'outEnd', 
                         labelkleur = 'black', 
                         ylimiet = 1) {
  
  chart_theme <- mschart_theme(main_title = fp_text(font.size = 9, bold = TRUE, font.family = lettertype, color = 'black'),
                               axis_title_x = fp_text(font.size = 9, font.family = lettertype),
                               axis_title_y = fp_text(font.size = 9, font.family = lettertype),
                               axis_text_x = fp_text(font.size = 8, font.family = lettertype),
                               axis_text_y = fp_text(font.size = 9, font.family = lettertype),
                               legend_text = fp_text(font.size = 9, font.family = lettertype),
                               grid_major_line = fp_border(style = "none"),
                               axis_ticks_x =  fp_border(width = 1),
                               axis_ticks_y =  fp_border(width = 1),
                               legend_position = legendapositie) 
  
  grafiek %>% set_theme(chart_theme) %>%
    chart_data_stroke('transparent') %>%
    chart_labels(title = grafiektitel, xlab = NULL, ylab = NULL) %>%
    chart_ax_x(major_tick_mark = 'none') %>%
    chart_ax_y(limit_min = 0, num_fmt = '0%', major_tick_mark = 'none', limit_max = ylimiet) %>%
    chart_data_labels(show_val = datalabels, num_fmt = '0%', position = labelpositie) %>%
    chart_labels_text(fp_text(font.size = 8, font.family = lettertype, color = labelkleur))
  
}

# Functie aanmaken om grafiekkleuren te bepalen
grafiekkleur <- function(cijfers, kleuren = c('#009898', '#91caca', '#b9b441', '#dcd9a6', '#e8525f', '#e9acb0'), type = 'standaard') {
  
  if(type == 'aslabel')
    
    kleuren[1:length(unique(cijfers$aslabel))] %>% set_names(unique(cijfers$aslabel))
  
  else if('groepering' %in% names(cijfers)) 
    
    kleuren[1:length(unique(cijfers$groepering))] %>% set_names(unique(cijfers$groepering))
  
  else if(any(!c('uitsplitsing', 'groepering') %in% names(cijfers))) 
    
    kleuren[c(1,5,3,2,6,4)][1:length(unique(cijfers$niveau))] %>% set_names(unique(cijfers$niveau))
  
  else kleuren[1]
  
}


# 4. Functie maken om cijfers te berekenen --------------------------------

# Functie aanmaken om cijfers te berekenen
bereken_cijfers <- function(data, omschrijving, type, indicator, waarde, indicator_label, uitsplitsing, groepering, niveau, jaar, 
                            weging_type, weegfactor_indicator, slidenummer, naam_aanduiding, niveau_waarde, niveau_naam, jaar_indicator,
                            Nvar = 30, Ncel = 5) {
  data %>%
    filter(.[[niveau]] == niveau_waarde & .[[jaar_indicator]] %in% jaar) %>%
    select(all_of(setdiff(c(jaar_indicator, indicator, uitsplitsing, groepering, weegfactor_indicator), NA))) %>%
    group_by(across(all_of(setdiff(c(jaar_indicator, indicator, uitsplitsing, groepering), NA)))) %>%
    {if(weging_type == 'geen') . else rename_at(., vars(matches(weegfactor_indicator)), ~ str_replace(., weegfactor_indicator, 'weegfactor'))} %>%
    summarise(n_cel = n(), 
              n_cel_gewogen = ifelse(weging_type == 'geen', NA, sum(weegfactor)), 
              .groups = 'drop') %>%
    drop_na(-n_cel_gewogen) %>%
    group_by(across(all_of(setdiff(c(jaar_indicator, uitsplitsing, groepering), NA)))) %>%
    mutate(n_totaal = sum(n_cel),
           n_totaal_gewogen = sum(n_cel_gewogen),
           n_min = n_totaal-n_cel,
           p = ifelse(test = n_totaal < Nvar | n_cel < Ncel | n_min < Ncel | n_cel == n_totaal, 
                      yes = NA, 
                      no = if(weging_type == 'geen') n_cel/n_totaal else n_cel_gewogen/n_totaal_gewogen) %>% as.numeric(),
           indicator = indicator,
           omschrijving = omschrijving,
           niveau = niveau_naam) %>%
    {if(is.na(waarde)) . else filter_at(., vars(all_of(indicator)), function(x) x %in% waarde)} %>%
    ungroup() %>%
    rename(jaar = 1, aslabel = 2) %>%
    {if(!is.na(indicator_label)) mutate(., aslabel = indicator_label) 
      else if(type == 'staafgrafiek gestapeld' | type == 'tabel') . 
      else mutate(., aslabel = omschrijving)}  %>%
    {if(!is.na(uitsplitsing)) rename_at(., vars(matches(uitsplitsing)), ~ str_replace(., uitsplitsing, 'uitsplitsing')) else .} %>%
    {if(!is.na(groepering)) rename_at(., vars(matches(groepering)), ~ str_replace(., groepering, 'groepering')) else .} %>%
    mutate(across(!n_cel & !n_cel_gewogen & !n_totaal & !n_totaal_gewogen & !n_min & !p & !jaar, to_character)) %>%
    select(indicator, omschrijving, jaar, niveau, aslabel, everything())
  
}


# 5. Functies maken om content te genereren -------------------------------

# functie voor het berekenen van percentages
type_percentage <- function(cijfers) {
  
  cijfers %>%
    mutate(p = round(p*100)) %>%
    select(p) %>%
    unlist() %>%
    unname() %>%
    {if(length(.) != 1) '-' 
      else if(is.na(.)) '-' 
      else paste0(., '%')}
  
}

# functie voor het plaatsen van percentages in tekst
type_percentage_tekst <- function(cijfers) {
  
  cijfers <- mutate(cijfers, p = ifelse(is.na(p), '-', paste0(round(p*100), '%')))
  
  reduce2(cijfers$indicator, cijfers$p,  .init = cijfers$omschrijving[1], str_replace)
  
}

# functie voor het genereren van een staafgrafiek
type_staafgrafiek <- function(cijfers, horizontaal = F, selectie = F, selectie_n = NA) {
  
    cijfers %>%
    {if(selectie == T) slice_max(., p, n = selectie_n, na_rm = T) else .} %>%
    {if(horizontaal == T) mutate(., across(intersect(names(.), c('niveau', 'aslabel', 'uitsplitsing', 'groepering')), fct_rev)) else .} %>%
    ms_barchart(x = if('uitsplitsing' %in% names(cijfers)) 'uitsplitsing' 
                    else 'aslabel',
                y = 'p',
                group = if('groepering' %in% names(cijfers)) 'groepering' 
                        else if(any(!c('uitsplitsing', 'groepering') %in% names(cijfers))) 'niveau' 
                        else NULL) %>%
    chart_data_fill(grafiekkleur(cijfers)) %>%
    grafiekstijl(grafiektitel = unique(cijfers$omschrijving),
                 legendapositie = if('groepering' %in% names(cijfers) | length(unique(cijfers$niveau)) > 1) 'b' else 'n') %>%
    {if(horizontaal == T) chart_settings(., dir = "horizontal") else .}
  
}

# functie voor het genereren van een staafgrafiek
type_staafgrafiek_gestapeld <- function(cijfers, horizontaal = F) {
  
  cijfers %>%
    {if(horizontaal == T) mutate(., across(intersect(names(.), c('niveau', 'uitsplitsing', 'groepering')), fct_rev)) else .} %>%
    ms_barchart(x = if('groepering' %in% names(cijfers)) 'groepering' else 'niveau', 
                y = 'p', 
                group = 'aslabel') %>%
    chart_data_fill(grafiekkleur(cijfers, type = 'aslabel')) %>%
    grafiekstijl(grafiektitel = unique(cijfers$omschrijving),
                 legendapositie = 'b',
                 labelpositie = 'ctr',
                 labelkleur = 'white') %>%
     as_bar_stack(dir = if(horizontaal == F) 'vertical' else 'horizontal')
  
}

# functie voor het genereren van een trendgrafiek
type_trendgrafiek <- function(cijfers, valuelabels = F) {
  
  cijfers %>%
    mutate(jaar = as.character(jaar)) %>%
    ms_linechart(x = 'jaar', y = 'p', group = 'niveau') %>%
    grafiekstijl(grafiektitel = unique(cijfers$omschrijving), datalabels = F, labelpositie = 't',
                 ylimiet = if(.$data %>% .$p %>% max(na.rm = T) < 0.10) 0.25 
                              else if(.$data %>% .$p %>% max(na.rm = T) < 0.25) 0.5 
                              else 1) %>%
    chart_data_fill(grafiekkleur(cijfers)) %>%
    chart_data_stroke(grafiekkleur(cijfers)) %>%
    chart_data_size(5) %>%
    {if(valuelabels == T) chart_data_labels(., position = 't', show_val = T, num_fmt = '0%') %>%
                          chart_labels_text(grafiekkleur(cijfers) %>% map(~fp_text(font.size = 8, color = ., font.family = 'Century Gothic')))
      else .}
}

# functie voor het genereren van een tabel
type_tabel <- function(cijfers) {
  
  tabel <- cijfers %>%
    select(niveau, aslabel, n_cel) %>%
    pivot_wider(names_from = niveau, values_from = n_cel) %>%
    rename(`Aantal ingevulde vragenlijsten` = aslabel)
  
  tabel %>%
    flextable() %>%
    border_remove() %>%
    fontsize(size = 12, part = "all") %>%
    font(fontname = 'Century Gothic', part = "all") %>%
    align(align = "center", part = "all") %>%
    bg(bg = '#e8525f', part = 'header') %>%
    bg(bg = '#009898', part = 'body') %>% 
    color(color = "white", part = "all") %>%
    border_outer(part="all", border = fp_border_default(width = 2, color = "white") ) %>%
    border_inner_h(border = fp_border_default(width = 2, color = "white"), part="all") %>%
    border_inner_v(border = fp_border_default(width = 2, color = "white"), part="all") %>%
    set_table_properties(layout = "fixed") %>%
    height_all(height = 1.03, unit = 'cm') %>%
    width(width = 5.85, unit = 'cm')
  
}


# 6. Functie aanmaken om content te genereren en te plaatsen --------------

# Functie aanmaken die op basis van de kolom 'type' uit de configuratie content genereert.
content_plaatsen <- function(data, rapportnaam, template, 
                             omschrijving, type, indicator, waarde, indicator_label, uitsplitsing, groepering, niveau, jaar,
                             weging_type, weegfactor_indicator, slidenummer, naam_aanduiding,
                             rapportconfiguratie) {
  
  cat(paste(rapportnaam, 'slidenummer:', slidenummer, 'indicator:', indicator))
  
  {
  
  if(type == 'rapportnaam') {
    
    value <- rapportnaam
    
  } else if(type == 'datum') {
    
    value <- Sys.Date() %>% format(format = "%d-%m-%Y") %>% as.character()
    
  } else 
    
    cijfers <- data.frame(omschrijving = omschrijving, type = type, indicator = indicator, waarde = waarde, indicator_label = indicator_label,
                          uitsplitsing = uitsplitsing, groepering = groepering, niveau = niveau, jaar = jaar,
                          weging_type = weging_type, weegfactor_indicator = weegfactor_indicator, 
                          slidenummer = slidenummer, naam_aanduiding = naam_aanduiding) %>%
      mutate(across(c(indicator, waarde, indicator_label, uitsplitsing, groepering, niveau, jaar), ~ .x %>% str_split(';\\s*'))) %>%
      {if(type != 'staafgrafiek gestapeld') unnest(., cols = c(indicator, indicator_label))
        else unnest(., cols = c(waarde, indicator_label))} %>%
      unnest(cols = c(uitsplitsing)) %>%
      unnest(cols = c(groepering)) %>%
      unnest(cols = c(niveau)) %>%
      mutate(weegfactor_indicator = ifelse(test = weging_type == 'niveau',
                                           yes = rapportconfiguratie$weegfactor_indicator[match(.$niveau, rapportconfiguratie$niveau_indicator)],
                                           no = weegfactor_indicator)) %>%
      cbind(rapportconfiguratie[match(.$niveau, rapportconfiguratie$niveau_indicator), -c(1,4)]) %>%
      pmap(bereken_cijfers, data = data) %>%
      reduce(bind_rows) %>%
      {if(length(unique(.$niveau)) > 1 & !is.na(groepering)) mutate(., groepering = paste(niveau, groepering)) else .} %>%
      mutate(across(intersect(names(.), c('niveau', 'aslabel', 'uitsplitsing', 'groepering')), ~factor(.x, levels = unique(.x), ordered = T)))
  
    
  if(type == 'percentage')

    value <- type_percentage(cijfers)
  
  else if(type == 'percentage in tekst')
    
    value <- type_percentage_tekst(cijfers)

  else if(type == 'staafgrafiek')

    value <- type_staafgrafiek(cijfers)

  else if(type == 'staafgrafiek liggend')

    value <- type_staafgrafiek(cijfers, horizontaal = T)
  
  else if(type == 'staafgrafiek gestapeld')
    
    value <- type_staafgrafiek_gestapeld(cijfers, horizontaal = T)
  
  else if(str_detect(type, '^top') == T)
    
    value <- type_staafgrafiek(cijfers, selectie = T, selectie_n = as.numeric(str_extract(type,'\\d+'))) 

  else if(type == 'trendgrafiek')
    
    value <- type_trendgrafiek(cijfers, valuelabels = T)
  
  else if(type == 'tabel')
    
    value <- type_tabel(cijfers)

  }

   on_slide(x = template, index = slidenummer) %>%
     ph_with(value = value, location = ph_location_label(ph_label = naam_aanduiding))

   cat("\014")
   
}


# 7. Rapportage maken -----------------------------------------------------

# Functie aanmaken om de rapportage te maken. 
rapportage_maken <- function(data, configuratie, rapportnaam, template, slideconfiguratie, 
                             niveau_indicator, niveau_waarde, niveau_naam, weegfactor_indicator, jaar_indicator) {
  
  
  template <- read_pptx(template)

  if('indeling' %in% names(configuratie[[slideconfiguratie]]) == T) {
  
    walk(.x = configuratie[[slideconfiguratie]] %>% select(indeling, slidenummer) %>% unique() %>% pull(indeling), 
         .f = add_slide, x = template, master = 'Figurenboek')
    
  }

  rapportconfiguratie <- data.frame(niveau_indicator = niveau_indicator,
                                    niveau_waarde = niveau_waarde,
                                    niveau_naam = niveau_naam,
                                    weegfactor_indicator = weegfactor_indicator,
                                    jaar_indicator = jaar_indicator) %>%
    mutate(across(c(niveau_indicator, niveau_waarde, niveau_naam, weegfactor_indicator), ~ .x %>% str_split(';\\s*'))) %>%
    unnest(cols = everything())
  
  slideconfiguratie <- configuratie[[slideconfiguratie]] %>%
    select(omschrijving, type, indicator, waarde, indicator_label, uitsplitsing, groepering, niveau, jaar,
           weging_type, weegfactor_indicator, slidenummer, naam_aanduiding)
  
  pmap(.l = slideconfiguratie,
       .f = content_plaatsen,
       data = data,
       rapportnaam = rapportnaam,
       template = template,
       rapportconfiguratie = rapportconfiguratie)
  
  print(template, paste0('Rapport ', rapportnaam, '.pptx'))
  
}


# 8. Rapportages maken op basis van rapportconfiguratie -------------------

# Configuratie inladen
configuratie <- padnaam_configuratie %>% 
  excel_sheets() %>% 
  set_names() %>% 
  map(read_excel, path = padnaam_configuratie)

# Rapportages aanmaken
pmap(.l = configuratie$Rapportconfiguratie, 
     .f = rapportage_maken, 
     data = data,
     configuratie = configuratie)


# Handmatig cijfers berekenen ter controle --------------------------------

# Dit deel van het script is optioneel en kan helpen bij het oplossen van errors
# Functie aanmaken om voor specifieke regels in de slideconfiguratie de cijfers te berekenen
handmatig_cijfers_berekenen <- function(configuratie, regel_rapportconfig, regel_slideconfig) {
  
  rapportconfiguratie <- configuratie$Rapportconfiguratie %>% 
    slice(regel_rapportconfig) %>%
    mutate(across(c(niveau_indicator, niveau_waarde, niveau_naam, weegfactor_indicator), ~ .x %>% str_split(';\\s*'))) %>%
    unnest(cols = everything())
  
  slideconfiguratie <- configuratie[[rapportconfiguratie$slideconfiguratie[1]]] %>%
    slice(regel_slideconfig)
   
  slideconfiguratie %>%
    select(omschrijving, type, indicator, waarde, indicator_label, uitsplitsing, groepering, niveau, jaar,
           weging_type, weegfactor_indicator, slidenummer, naam_aanduiding) %>%
    mutate(across(c(indicator, waarde, indicator_label, uitsplitsing, groepering, niveau, jaar), ~ .x %>% str_split(';\\s*'))) %>%
    {if(slideconfiguratie$type != 'staafgrafiek gestapeld') unnest(., cols = c(indicator, indicator_label))
      else unnest(., cols = c(waarde, indicator_label))} %>%
    unnest(cols = c(uitsplitsing)) %>%
    unnest(cols = c(groepering)) %>%
    unnest(cols = c(niveau)) %>%
    mutate(weegfactor_indicator = ifelse(test = weging_type == 'niveau',
                                         yes = rapportconfiguratie$weegfactor_indicator[match(.$niveau, rapportconfiguratie$niveau_indicator)],
                                         no = weegfactor_indicator)) %>%
    cbind(rapportconfiguratie[match(.$niveau, rapportconfiguratie$niveau_indicator), -c(1:4,7)]) %>%
    pmap(bereken_cijfers, data = data) %>%
    reduce(bind_rows) %>%
    {if(length(unique(.$niveau)) > 1 & !is.na(slideconfiguratie$groepering)) mutate(., groepering = paste(niveau, groepering)) else .} %>%
    mutate(across(intersect(names(.), c('niveau', 'aslabel', 'uitsplitsing', 'groepering')), ~factor(.x, levels = unique(.x), ordered = T)))
  
}

handmatig_cijfers_berekenen(configuratie = configuratie, 
                            regel_rapportconfig = 1,
                            regel_slideconfig = 5)
