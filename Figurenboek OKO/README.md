* Wil je gewogen (en/of ongewogen) cijfers in je figurenboek dan kun je gebruikmaken van de bestanden:
  * '1. Configuratie Figurenboek.xlsx'
  * '2. Template Figurenboek.pptx'
  * 'Rapportage.R' (Het script staat in de hoofdmap: https://github.com/ggdatascience/rapportage_monitor_gmj/blob/main/Rapportage.R)

_Stijltip:_ In het figurenboek is het mooier als de legenda boven aan de staafgrafiek staat. Wil je dit ook? Pas dan regel 179 van het script aan waarbij je de 'b' verandert in 't'.

originieel: 
legendapositie = if('groepering' %in% names(cijfers) | length(unique(cijfers$niveau)) > 1) **'b'** else 'n'

aangepast: 
legendapositie = if('groepering' %in% names(cijfers) | length(unique(cijfers$niveau)) > 1) **'t'** else 'n'

* Versie 1 van het figurenboek is ook nog beschikbaar, maar omdat deze geen nieuwe versies meer krijgt is het verstandig om deze niet meer te gebruiken. Heb je hier al eerder gebruik van gemaakt en wil je dit script blijven gebruiken (voor ongewogen cijfers), gebruik dan de volgende bestanden:
  *  '1. Configuratie Figurenboek Ongewogen.xlsx'
  *  '2. Template Figurenboek.pptx'
  *  '3. Script Figurenboek Ongewogen.R'


