# rapportage_monitor_gmj
Script om schoolrapportages voor de Gezondheidsmonitor Jeugd te maken in PowerPoint op basis van een configuratiebestand in Excel.

Tip om een regiobeeld te genereren:
- Voeg in het Excel configuratiebestand de schoolcode "totaal" toe met als referentie totaal.
- In het rapportage script voeg je dit toe: ***data$totaal <- "totaal"***

Tip om trends te maken, data van andere jaren in te laden (op voorwaarde dat de indicatoren dezelfde naam hebben en er een jaarvariabele inzit)
- Deze code is te gebruiken om 2021 data in te lezen:
***data2021 <- read_spss('Trendbestanden/Trendbestand 2021 met indicatoren 2023 RVO-scholen.sav')) %>%
  select(indicatoren2021) # Selecteer alleen indicatoren die in databestand van 2021 aanwezig moeten zijn.***
  ***data2019 <- read_spss('Trendbestanden/Trendbestand 2019 met indicatoren 2023 RVO-scholen.sav')) %>%
  select(indicatoren2019) # Selecteer alleen indicatoren die in databestand van 2019 aanwezig moeten zijn.***
- Vervolgens de data samenvoegen tot een volledige trend dataset:
  ***data <- full_join(data2023, data2021) %>% 
  full_join(data2019)***

Tip om foutmelding op basis van labels 0, 8 of 9 te voorkomen
- Zorg ervoor dat de labels 0, 8 en 9 omgezet worden naar NULL:
- ***# Value labels 0, 8 en 9 eruit halen
val_label(data$KLAS, 0) <- NULL
val_label(data$KLAS, 9) <- NULL
val_label(data$GENDER, 0) <- NULL
val_label(data$GENDER, 9) <- NULL
val_label(data$MBOKK3S31, 8) <- NULL
val_label(data$MBOKK3S31, 9) <- NULL
val_label(data$MBOKA3S31, 9) <- NULL
val_label(data$MBGSK3S2, 9) <- NULL
val_label(data$SBOSK3S7, 9) <- NULL***
