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
data <- read_spss('Testdata GMJ2023.sav')

# Indicator onderwijsniveau hercoderen naar factor variabele (anders wordt deze in alfabetische volgorde getoond in figuren)
data$MBOKK3S31 <- factor(data$MBOKK3S31, levels = c('Vmbo', 'Havo/Vwo'), ordered = T)

# 2. Excelbestand met configuratie laden ----------------------------------

# Padnaam opgeven van de Excel configuratie. Staat dit bestand in je working directory dan hoef je alleen de bestandsnaam
# op te geven (inclusief de bestandsextensie)
padnaam_configuratie <- '1. Configuratie.xlsx'

# Inladen van alle tabbladen van de Excel configuratie
configuratie <- padnaam_configuratie %>% 
  excel_sheets() %>% 
  set_names() %>% 
  map(read_excel, path = padnaam_configuratie)


# 3. Grafiekinstellingen en opmaak ----------------------------------------

# Kleurpalet aanmaken
kleuren <- c('#009898', '#91caca', '#b9b441', '#dcd9a6', '#e8525f', '#e9acb0')

# Lettertype importeren. In de console wordt gevraagd of je dit wil doen run 'y' in de console om te importeren.
extrafont::font_import(pattern = 'GOTHIC')

# Check om te zien of het lettertype goed is ingeladen (check of het lettertype Century Gothic ertussen staat)
windowsFonts()

# Lettertype bepalen
lettertype <- "Century Gothic"

# Thema voor staafgrafiek aanmaken
chart_theme <- mschart_theme(
  main_title = fp_text(font.size = 9, bold = TRUE, font.family = lettertype, color = 'black'),
  axis_title_x = fp_text(font.size = 9, font.family = lettertype),
  axis_title_y = fp_text(font.size = 9, font.family = lettertype),
  axis_text_x = fp_text(font.size = 8, font.family = lettertype),
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
  
  result <- data %>%
    filter(.data[[niveau_var]] == niveau_label & .data[[var_jaar]] %in% jaar) %>%
    select(all_of(setdiff(c(var_jaar, indicator, uitsplitsing, groepering), NA))) %>%
    group_by(across(all_of(setdiff(c(var_jaar, indicator, uitsplitsing, groepering), NA))), .drop = T) %>%
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
  
  # Perform chi-squared test
  
  if(toetsen == T){
    
    chi_squared_result <- chisq.test(result[, c("n", "nmin")])
    
    hoogste_waarde <- which.max(result$val)
    rij_met_hoogste_waarde <- result[hoogste_waarde, ]
    laagste_waarde <- which.min(result$val)
    rij_met_laagste_waarde <- result[laagste_waarde, ]
    
    # Extract 'omschrijving' for highest and lowest values
    hoogste_omschrijving <- rij_met_hoogste_waarde$omschrijving
    laagste_omschrijving <- rij_met_laagste_waarde$omschrijving
    hoogste_categorie <- rij_met_hoogste_waarde[2]
    laagste_categorie <- rij_met_laagste_waarde[2]
    
    # Add chi-squared test results to the result dataframe
    result$p_val <- chi_squared_result$p.value
    result$chi_val <- chi_squared_result$statistic
    result$df <- chi_squared_result$parameter
    
  }
  
  return(result)
  
  
}

# bereken_cijfers(data = data,
#                 basis = 'Schoolcode', basis_label = 'Yuverta Roermond', referentie = 'totaal', referentie_label = 'totaal',
#                 omschrijving = 'Mening ouders alcohol', indicator = 'LOAGK316', waarden = c(1,2,3,4,5,6), valuelabel = NA, uitsplitsing = NA, groepering = NA, niveau = 'basis', jaar = 2023, 
#                 var_jaar = 'Onderzoeksjaar', Nvar = 30, Ncel = 5, toetsen = F)


# 5. Functies maken om content te genereren -------------------------------

# functie voor het berekenen van percentages
type_percentage <- function(data, 
                            basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                            indicator, waarden, niveau, jaar) {
  data %>%
    bereken_cijfers(basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                    indicator = indicator, waarden = waarden, niveau = niveau, jaar = jaar) %>%
    mutate(., val = round(val*100)) %>%
    select(val) %>%
    unlist() %>%
    unname() %>%
    {if(length(.) != 1) '-' 
      else if(is.na(.)) '-' 
      else paste0(., '%')}
      
}

type_percentage_tekst <- function(data, basis = NA, basis_label = NA, valuelabel = NA, omschrijving = NA, uitsplitsing = NA, groepering = NA, referentie = NA, referentie_label = NA, indicator, waarden, niveau, jaar, format = "percentage tekst") {
  
  data.frame(indicator = indicator, valuelabel = valuelabel) %>%
    pmap(bereken_cijfers, data = data, basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label, waarden = waarden, niveau = niveau, jaar = jaar, omschrijving = omschrijving, uitsplitsing = uitsplitsing, groepering = groepering) %>% reduce(bind_rows) %>% data.frame(., y = "val") %>% select(c(indicator, val, omschrijving)) %>%
    {if(as.numeric(length(indicator)) == 1) mutate(., val = round(val*100)) %>% pivot_wider(names_from = indicator, values_from = val) %>% 
        {if(is.na(.[[2]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[2], paste0("-"))) %>% select(omschrijving) %>% unlist() %>% unname()
          else if(!is.na(.[[2]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[2], paste0(.[[2]], "%"))) %>% select(omschrijving) %>% unlist() %>% unname()}
      else if(as.numeric(length(indicator)) == 2) mutate(., val = round(val*100)) %>% pivot_wider(names_from = indicator, values_from = val) %>%
        {if(is.na(.[[2]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[2], paste0("-")))
          else if(!is.na(.[[2]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[2], paste0(.[[2]], "%")))
        } %>% 
        {if(is.na(.[[3]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[3], paste0("-"))) %>% select(omschrijving) %>% unlist() %>% unname()
          else if(!is.na(.[[3]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[3], paste0(.[[3]], "%"))) %>% select(omschrijving) %>% unlist() %>% unname()
        }
      else if(as.numeric(length(indicator)) == 3) mutate(., val = round(val*100)) %>% pivot_wider(names_from = indicator, values_from = val) %>%
        {if(is.na(.[[2]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[2], paste0("-")))
          else if(!is.na(.[[2]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[2], paste0(.[[2]], "%")))
        } %>% 
        {if(is.na(.[[3]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[3], paste0("-")))
          else if(!is.na(.[[3]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[3], paste0(.[[3]], "%")))
        } %>% 
        {if(is.na(.[[4]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[4], paste0("-"))) %>% select(omschrijving) %>% unlist() %>% unname()
          else if(!is.na(.[[4]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[4], paste0(.[[4]], "%"))) %>% select(omschrijving) %>% unlist() %>% unname()
        }
      else if(as.numeric(length(indicator)) == 4) mutate(., val = round(val*100)) %>% pivot_wider(names_from = indicator, values_from = val) %>%
        {if(is.na(.[[2]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[2], paste0("-")))
          else if(!is.na(.[[2]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[2], paste0(.[[2]], "%")))
        } %>% 
        {if(is.na(.[[3]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[3], paste0("-")))
          else if(!is.na(.[[3]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[3], paste0(.[[3]], "%")))
        } %>% 
        {if(is.na(.[[4]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[4], paste0("-")))
          else if(!is.na(.[[4]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[4], paste0(.[[4]], "%")))
        } %>% 
        {if(is.na(.[[5]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[5], paste0("-"))) %>% select(omschrijving) %>% unlist() %>% unname()
          else if(!is.na(.[[5]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[5], paste0(.[[5]], "%"))) %>% select(omschrijving) %>% unlist() %>% unname()
        }
      else if(as.numeric(length(indicator)) == 5) mutate(., val = round(val*100)) %>% pivot_wider(names_from = indicator, values_from = val) %>% 
        {if(is.na(.[[2]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[2], paste0("-")))
          else if(!is.na(.[[2]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[2], paste0(.[[2]], "%")))
        } %>% 
        {if(is.na(.[[3]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[3], paste0("-")))
          else if(!is.na(.[[3]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[3], paste0(.[[3]], "%")))
        } %>% 
        {if(is.na(.[[4]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[4], paste0("-")))
          else if(!is.na(.[[4]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[4], paste0(.[[4]], "%")))
        } %>%
        {if(is.na(.[[5]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[5], paste0("-")))
          else if(!is.na(.[[5]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[5], paste0(.[[5]], "%")))
        } %>% 
        {if(is.na(.[[6]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[6], paste0("-"))) %>% select(omschrijving) %>% unlist() %>% unname()
          else if(!is.na(.[[6]])) data.frame(.) %>% mutate(omschrijving = str_replace(omschrijving, colnames(.)[6], paste0(.[[6]], "%"))) %>% select(omschrijving) %>% unlist() %>% unname()
        }
    }
}

# type_percentage_tekst(data, basis = 'Schoolcode', basis_label = 'College van de Hoge Hoed',
#                       omschrijving = 'LBSDK3S1 van de jongeren op uw school heeft ooit wiet of hasj gebruikt. Het percentage dat in de laatste 4 weken wiet of hasj gebruikt heeft is LBSDK3S2.',
#                       referentie = 'totaal', referentie_label = 'totaal', indicator = c("LBSDK3S1", "LBSDK3S2"), waarden = 1, niveau ='basis', jaar = 2023)

# functie voor het vergelijken van twee groepen
type_vergelijking <- function(data, 
                              basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                              indicator, uitsplitsing, groepering, omschrijving, waarden, niveau, jaar, toetsen = T) {
  
  cijfers <-data %>%
    bereken_cijfers(basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                    indicator = indicator, waarden = waarden, niveau = niveau, jaar = jaar, omschrijving = omschrijving, 
                    uitsplitsing = uitsplitsing, groepering = groepering, toetsen = toetsen) %>%
    mutate(percentage = paste0(round(val*100, 0), '%')) %>%
    select(-indicator, -jaar, -niveau, -aslabel, -n, -ntot, -val, -nmin, -chi_val, -df) %>%
    rename(labels = 2)  %>%
    pivot_wider(names_from = labels,
                values_from = percentage) 
  
  # Volgorde van kolommen aanpassen zodat de eerstgenoemde groep in de tekst ook de eerste groep (derde kolom) in het dataframe is 
  # en de laatste genoemde groep in de tekst de tweede groep (vierde kolom)
  {if((cijfers$omschrijving %>% str_locate(names(cijfers)[3]) %>% .[1]) > (cijfers$omschrijving %>% str_locate(names(cijfers)[4]) %>% .[1])) select(cijfers, 1,2,4,3)
  else cijfers} %>%
    mutate(omschrijving = str_replace(omschrijving, paste0('\\[', names(.)[3], '\\]'), unlist(.[1,3])),
           omschrijving = str_replace(omschrijving, paste0('\\[', names(.)[4], '\\]'), unlist(.[1,4])),
           omschrijving = {if(p_val >= 0.05) str_replace(omschrijving, '\\[hoger dan/lager dan/vergelijkbaar met\\]', 'vergelijkbaar met')
                           else if(p_val < 0.05) str_replace(omschrijving, '\\[hoger dan/lager dan/vergelijkbaar met\\]',
                                                             ifelse(unlist(.[1,3]) > unlist(.[1,4]), 'hoger dan', 'lager dan'))}) %>%
    pull(omschrijving) 
  
}

# type_vergelijking(data = data, toetsen = T,
#                   basis = 'Schoolcode', basis_label = 'College van de Hoge Hoed', referentie = 'totaal', referentie_label = 'totaal',
#                   omschrijving = 'Van de meisjes voelt [Meisje] zich gelukkig en dit percentage is [hoger dan/lager dan/vergelijkbaar met] jongens, waarvan [Jongen] zich gelukkig voelt.',
#                   indicator = 'EBEGK3S2', waarden = 1, uitsplitsing = 'GENDER', groepering = NA, niveau = 'basis', jaar = 2023)

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
          grafiekstijl(grafiektitel = omschrijving, 
                       ylimiet = ifelse((.$data %>% .$val %>% max(na.rm = T)) < 0.5, 0.5, 1))
  
      else if(!is.na(uitsplitsing) & is.na(groepering))
          ms_barchart(., x = 'uitsplitsing', y = 'val') %>%
          chart_data_fill(kleuren[1]) %>%
          set_theme(chart_theme) %>%
          chart_theme(legend_position = 'n') %>%
          grafiekstijl(grafiektitel = omschrijving, 
                     ylimiet = ifelse((.$data %>% .$val %>% max(na.rm = T)) < 0.5, 0.5, 1))
  
      else if(is.na(uitsplitsing) & !is.na(groepering))
          ms_barchart(., x = 'omschrijving', y = 'val', group = 'groepering') %>%
          chart_data_fill(kleuren[1:(val_labels(data[[groepering]]) %>% length())] %>% set_names(val_labels(data[[groepering]]) %>% names())) %>%
          set_theme(chart_theme) %>%
          chart_ax_x(tick_label_pos = 'none') %>%
          grafiekstijl(grafiektitel = omschrijving, 
                     ylimiet = ifelse((.$data %>% .$val %>% max(na.rm = T)) < 0.5, 0.5, 1))
  
      else
        ms_barchart(., x = 'omschrijving', y = 'val') %>%
          chart_data_fill(kleuren[1]) %>%
          set_theme(chart_theme) %>%
          chart_theme(legend_position = 'n') %>%
          chart_ax_x(tick_label_pos = 'none') %>%
          grafiekstijl(grafiektitel = omschrijving, 
                       ylimiet = ifelse((.$data %>% .$val %>% max(na.rm = T)) < 0.5, 0.5, 1))
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
      ms_barchart(x = 'aslabel', y = 'val', group = 'niveau') %>% # x is nu 'aslabel' en niet 'omschrijving' kan mogelijk problemen opleveren
      chart_data_fill(kleuren[c(1,5)] %>% set_names(c('School', 'Regio'))) %>%
      set_theme(chart_theme) %>%
      chart_theme(legend_position = 'b') %>%
      chart_ax_x(tick_label_pos = 'none') %>%
      chart_settings(overlap = -20) %>%
      grafiekstijl(grafiektitel = omschrijving,
                   ylimiet = ifelse((.$data %>% .$val %>% max(na.rm = T)) < 0.5, 0.5, 1))
    
  }
}

# type_staafgrafiek(data = data,
#                   basis = 'Schoolcode', basis_label = 'College van de Hoge Hoed', referentie = 'totaal', referentie_label = 'totaal',
#                   omschrijving = 'grafiektitel', indicator = 'LOMVK301', waarden = 1, uitsplitsing = NA, groepering = 'KLAS', niveau = 'basis', jaar = 2023) %>%
#   print(preview = T)

# functie voor gestapelde staafgrafiek
type_staafgrafiek_gestapeld <- function(data,
                                        basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                                        omschrijving = NA, indicator, waarden = NA, niveau, jaar,
                                        valuelabel = NA, horizontaal = F) {
  
  if(niveau == 'basis' | niveau == 'referentie') {
    
    cijfers <- bereken_cijfers(data = data, basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                               indicator = indicator, waarden = waarden, niveau = niveau, jaar = jaar, omschrijving = omschrijving, valuelabel = valuelabel) %>%
               mutate(aslabel = factor(aslabel, levels = unique(aslabel), ordered = T))
    
    legendakleuren <- kleuren[1:length(cijfers$aslabel)] %>% set_names(cijfers$aslabel)
    
    ms_barchart(cijfers, x = 'niveau', y = 'val', group = 'aslabel') %>% 
      {if(horizontaal == T) as_bar_stack(., dir = 'horizontal') else .} %>%
      chart_settings(grouping = "stacked", gap_width = 50, overlap = 100 ) %>%
      chart_data_fill(legendakleuren) %>%
      set_theme(chart_theme) %>%
      chart_theme(legend_position = 'b') %>%
      grafiekstijl(grafiektitel = omschrijving,
                   ylimiet = 1,
                   labelpositie = 'ctr',
                   labelkleur = 'white')
    
  } else if(niveau == 'basis en referentie') {
    
    cijfers <- data.frame(niveau = c('basis', 'referentie'),
                          basis = c(basis, NA),
                          basis_label = c(basis_label, NA),
                          referentie = c(NA, referentie),
                          referentie_label = c(NA, referentie_label)) %>%
      pmap(bereken_cijfers, data = data, indicator = indicator, waarden = waarden, jaar = jaar,
                    omschrijving = omschrijving) %>%
      reduce(bind_rows) %>%
      mutate(niveau = factor(niveau, levels = c('School', 'Regio'), ordered = T),
             aslabel = factor(aslabel, levels = unique(aslabel), ordered = T))
    
    legendakleuren <- kleuren[1:length(unique(cijfers$aslabel))] %>% set_names(unique(cijfers$aslabel))
    
    ms_barchart(cijfers, x = 'niveau', y = 'val', group = 'aslabel') %>%
      {if(horizontaal == T) as_bar_stack(., dir = 'horizontal') else .} %>%
      chart_settings(grouping = "stacked", gap_width = 50, overlap = 100 ) %>%
      chart_data_fill(legendakleuren) %>%
      set_theme(chart_theme) %>%
      chart_theme(legend_position = 'b') %>%
      grafiekstijl(grafiektitel = omschrijving,
                   ylimiet = 1,
                   labelpositie = 'ctr',
                   labelkleur = 'white')
    
  }
}

# type_staafgrafiek_gestapeld(data = data, basis = 'Schoolcode', basis_label = 'Yuverta Roermond',
#                             referentie = 'totaal', referentie_label = 'totaal', omschrijving = 'Mening ouders alcohol',
#                             indicator = 'LOAGK316', waarden = c(1,2,3,4,5,6),
#                             niveau = 'basis en referentie', jaar = 2023, horizontaal = T) %>%
#   print(preview = T)

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
    grafiekstijl(grafiektitel = omschrijving, datalabels = F, labelpositie = 't',
                 ylimiet = ifelse((.$data %>% .$val %>% max(na.rm = T)) < 0.5, 0.5, 1)) %>%
    chart_data_fill(kleuren[c(1,5)] %>% set_names(c('School', 'Regio'))) %>%
    chart_data_stroke(kleuren[c(1,5)] %>% set_names(c('School', 'Regio'))) %>%
    chart_data_symbol('none')
}

# functie voor combigrafiek
type_combi <- function(data,
                       basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                       omschrijving = NA, indicator, waarden = 1, uitsplitsing = NA, groepering = NA, niveau, jaar,
                       valuelabel = NA, horizontaal = F, selectie = F, selectie_n = NA) {
  
  data.frame(indicator = indicator,
             valuelabel = valuelabel) %>%
    pmap(bereken_cijfers, data = data, basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
      waarden = waarden, niveau = niveau, jaar = jaar, omschrijving = omschrijving, uitsplitsing = uitsplitsing, groepering = groepering) %>%
    reduce(bind_rows) %>%
    {if(horizontaal == T & !is.na(groepering)) mutate(., groepering = fct_rev(groepering)) else .} %>% 
    {if(selectie == T) arrange(., desc(val)) %>%
      top_n(n = selectie_n) else .} %>%
    {if(!is.na(groepering))
      ms_barchart(.,x = 'aslabel', y = 'val', group = 'groepering') %>%
        chart_data_fill(kleuren[1:(val_labels(data[[groepering]]) %>% length())] %>% set_names(val_labels(data[[groepering]]) %>% names())) %>%
        set_theme(chart_theme)

      else ms_barchart(.,x = 'aslabel', y = 'val') %>%
        chart_data_fill(kleuren[1]) %>%
        set_theme(chart_theme) %>%
        chart_theme(legend_position = 'n')} %>%
    grafiekstijl(grafiektitel = omschrijving, ylimiet = ifelse((.$data %>% .$val %>% max(na.rm = T)) < 0.5, 0.5, 1)) %>%
    {if(horizontaal == T) chart_settings(., dir = "horizontal") else .}

}


# type_combi(data = data, basis = 'Schoolcode', basis_label = 'College van de Hoge Hoed',
#            referentie = NA, referentie_label = NA, omschrijving = 'Jongeren komen aan alcohol via',
#            indicator = c('LOAGK311', 'LOAGK312', 'LOAGK313', 'LOAGK314', 'LOAGK315'),
#            niveau = 'basis', jaar = 2023, groepering = 'KLAS',
#            valuelabel = c('Zelf kopen', 'Anderen kopen', 'Vrienden', 'Ouders', 'Andere volwassenen'), horizontaal = T) %>%
#   print(preview = T)

# functie voor responstabel
type_tabel <- function(data,
                       basis = NA, basis_label = NA,
                       indicator, var_jaar = 'Onderzoeksjaar', jaar = 2023,
                       valuelabel) {
  
  data %>%
    filter(.data[[basis]] == basis_label & .data[[var_jaar]] %in% jaar) %>%
    select(all_of(setdiff(c(indicator), NA))) %>%
    group_by(across(all_of(setdiff(c(indicator), NA)))) %>%
    tally() %>%
    drop_na() %>%
    ungroup()%>%
    mutate(across(!n, to_character)) %>%
    `colnames<-`(c(valuelabel, 'Aantal')) %>%
    flextable() %>%
    border_remove() %>%
    fontsize(size = 12, part = "all") %>%
    font(fontname = lettertype, part = "all") %>%
    align(align = "center", part = "all") %>%
    bg(bg = kleuren[5], part = "header") %>%
    bg(bg = kleuren[2], i = {if(.$body %>% .$data %>% nrow == 8) 1:4 else 1:2}) %>%
    bg(bg = kleuren[1], i = {if(.$body %>% .$data %>% nrow == 8) 5:8 else 3:4}) %>%
    color(color = "white", part = "all") %>%
    border_outer(part="all", border = fp_border_default(width = 1, color = "white") ) %>%
    border_inner_h(border = fp_border_default(width = 1, color = "white"), part="all") %>%
    border_inner_v(border = fp_border_default(width = 1, color = "white"), part="all") %>%
    set_table_properties(layout = "fixed") %>%
    height_all(height = 1.03, unit = 'cm') %>%
    width(width = 5.85, unit = 'cm')
  
}

# type_tabel(data = data, basis = 'Schoolcode', basis_label = 'Blariacum Venlo',
#            indicator = c('KLAS', 'MBOKK3S31', 'GENDER'), valuelabel = c('Leerjaar', 'Onderwijsniveau', 'Gender'))


# 6. Functie aanmaken om content te genereren en te plaatsen --------------

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
                             indicator = indicator, waarden = waarden, niveau = niveau, jaar = jaar)
  
  } else if(type == 'percentage in tekst') {
    
    value <- type_percentage_tekst(data = data, basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                                   omschrijving = omschrijving, indicator = indicator, waarden = waarden, niveau = niveau, jaar = jaar)
    
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
                        basis = basis, basis_label = basis_label, valuelabel = valuelabel, referentie = referentie, referentie_label = referentie_label,
                        omschrijving = omschrijving, indicator = indicator, uitsplitsing = uitsplitsing, groepering = groepering, waarden = waarden, niveau = niveau, jaar = jaar) 
  
  } else if(type == 'combi liggend') {
    
    value <- type_combi(data = data, horizontaal = T,
                        basis = basis, basis_label = basis_label, valuelabel = valuelabel, referentie = referentie, referentie_label = referentie_label,
                        omschrijving = omschrijving, indicator = indicator, uitsplitsing = uitsplitsing, groepering = groepering, waarden = waarden, niveau = niveau, jaar = jaar) 
    
  } else if(type == 'vergelijking') {
    
    value <- type_vergelijking(data = data, basis_label = basis_label, indicator = indicator, 
                      omschrijving = omschrijving, uitsplitsing = uitsplitsing)
  
  } else if(type == 'staafgrafiek gestapeld') {
    
    value <- type_staafgrafiek_gestapeld(data = data, basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                                         niveau = niveau, jaar = jaar, waarden = waarden,
                                         indicator = indicator, omschrijving = omschrijving, horizontaal = T)
    
  } else if(str_detect(type, '^top') == T) {
    
    value <- type_combi(data = data,
                        basis = basis, basis_label = basis_label, valuelabel = valuelabel, referentie = referentie, referentie_label = referentie_label,
                        omschrijving = omschrijving, indicator = indicator, uitsplitsing = uitsplitsing, groepering = groepering, waarden = waarden, niveau = niveau, jaar = jaar,
                        selectie = T, selectie_n = str_extract(type,'\\d+')) 
    
  } else if(type == 'tabel') {
    
    value <- type_tabel(data = data,
                        basis = basis, basis_label = basis_label,
                        indicator = indicator, jaar = jaar,
                        valuelabel = valuelabel)
    
  } 
  
  on_slide(template, index = index) %>%
    ph_with(value = value, location = ph_location_label(ph_label = label))
  
}


# 7. Rapportage maken -----------------------------------------------------

# Functie aanmaken om de rapportage te maken. 
rapportage_maken <- function(data, template, configuratie, basis, basis_label, referentie, referentie_label, slideconfiguratie, filteren = F) {
  
  slideconfiguratie <- configuratie[[slideconfiguratie]] %>%
    {if(length(filteren) == 1 & !is.numeric(filteren)) . else slice(., -filteren)} %>%
    select(-vraag) %>%
    mutate(across(c(indicator, waarden, valuelabel, jaar), ~ .x %>% str_split(';\\s*')))
  
  template <- read_pptx(template)
  pwalk(.l = slideconfiguratie, .f = content_plaatsen, data = data, template = template, basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label)
  print(template, paste0('Rapport ', basis_label, '.pptx'))

}

# 8. Rapportages maken op basis van rapportconfiguratie -------------------

pmap(.l = configuratie$Rapportconfiguratie, 
     .f = rapportage_maken, data = data, configuratie = configuratie, template = '2. Template.pptx')


# Foutmeldingen oplossen --------------------------------------------------

# Functie aanmaken om foutmeldingen op te lossen
foutmeldingen <- function(rapportconfiguratie, rapportregel, slideconfiguratie, slideregel, nieuw_venster = T) {
  
  rapport <- configuratie[[rapportconfiguratie]] %>%
    slice(rapportregel) %>%
    select(-slideconfiguratie)
  
  slide <- configuratie[[slideconfiguratie]] %>%
    select(-vraag) %>%
    mutate(across(c(indicator, waarden, valuelabel, jaar), ~ .x %>% str_split(';\\s*'))) %>%
    slice(slideregel)
  
  cbind(rapport, slide) %>%
    {if(nieuw_venster == T) View(., title = 'Foutmelding') else .}

}

# Foutmeldingen opsporen
foutmeldingen(rapportconfiguratie = 'Rapportconfiguratie',
              rapportregel = 1,
              slideconfiguratie = 'Slideconfiguratie Testdata',
              slideregel = 33)
