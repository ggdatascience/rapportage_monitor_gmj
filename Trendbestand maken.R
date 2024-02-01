# Data 2023 laden
data2023 <- read_spss('Scholenbestand GMJ 2023.sav') %>%
    mutate(Onderzoeksjaar = 2023)

# Data 2021 laden, vul bij de select() functie de variabelen in voor jouw gewenste trends en vergeet niet de schoolnaamvariabele toe te voegen
data2021 <- read_spss('Scholenbestand GMJ 2021.sav') %>%
  select(SCHOOLNAAMVARIABELE, REGIONAAMVARIABELE, EBEGK3S2, EBGLK3S1, LBAGK3S20, LBRAK3S19, PBMHK3S3, SBEZK3S2, SBRLK3S21, LBLBK3S13) %>%
  mutate(Onderzoeksjaar = 2021)

# Data 2019 laden, vul bij de select() functie de variabelen in voor jouw gewenste trends en vergeet niet de schoolnaamvariabele toe te voegen
data2019 <- read_spss('Scholenbestand GMJ 2019.sav') %>%
  select(SCHOOLNAAMVARIABELE, REGIONAAMVARIABELE, EBEGK3S2, EBGLK3S1, LBAGK3S20, LBRAK3S19, SBRLK3S21, LBLBK3S13) %>%
  mutate(Onderzoeksjaar = 2019)

# Bestanden samenvoegen
data <- data2023 %>%
  full_join(data2021) %>%
  full_join(data2019)


