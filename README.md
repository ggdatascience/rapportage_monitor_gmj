# Wijzigingen in Rapportage.R
## 22 maart 2024

In de meest recente update van het script is de mogelijkheid om te wegen toegevoegd. Het is ook mogelijk om met dit nieuwe script ongewogen cijfers te berekenen. Om te bepalen of en hoe er moet worden gewogen zijn aanpassingen gedaan aan de configuratie. Hieronder staan de aanpassingen weergegeven.

* **Drie opties voor weging toegevoegd aan het script**
  * **geen**: cijfers worden berekend zonder weging
  * **niveau**: weegfactor wordt bepaald op basis van de niveau_indicator
  * **indicator**: weegfactor wordt bepaald op basis van de indicator (dit is nodig voor de Gezondheidmonitor Volwassenen en Ouderen)
* **Rapportconfiguratie**: kolommen toegevoegd en naamgeving aangepast omdat de originele naamgeving verwarring veroorzaakte
  * Kolom _rapportnaam_ toegevoegd waarin kan worden aangegeven welke naam het rapport moet krijgen (deze kolom wordt gebruikt door als in de slideconfiguratie in de kolom _type_ rapportnaam staat aangegeven).
  * Kolom _template_ toegevoegd waarin de padnaam van het template moet worden opgegeven. Hierdoor is nu mogelijk om voor verschillende rapporten binnen dezelfde configuratie verschillende templates te gebruiken.
  * Kolom _slideconfiguratie_ is verplaatst naar de 3e kolom, maar is verder ongewijzigd.
  * Kolom _niveau_indicator_ toegevoegd waarin de indicatornamen moeten staan van de niveaus die gebruikt worden in het rapport. Als er meer dan een niveau wordt gebruikt dan moeten de indicatornamen gescheiden zijn door een puntkomma ';'.
  * Kolom _niveau_waarde_ toegevoegd waarin de waarde die hoort bij de opgegeven niveau_indicator moet worden opgegeven. Deze waarde kan zowel numeric als een character string zijn. Als er meer dan een niveau wordt gebruikt dan moeten de niveau waardes gescheiden zijn door een puntkomma ';'.
  * Kolom _niveau_naam_ toegevoegd waarin de naam kan worden bepaald die (per opgegeven niveau) in de figuren moet worden gebruikt. Als er meer dan een niveau wordt gebruikt dan moeten de niveau namen gescheiden zijn door een puntkomma ';'.
  * Kolom _weegfactor_indicator_ toegevoegd waarin per niveau aangegeven kan worden welke indicator voor de weegfactor moet worden gebruikt. Deze kolom hoeft alleen ingevuld te worden als er gewogen wordt op niveau. Voor ongewogen cijfers of voor weging op indicatorniveau kan deze kolom leegblijven. Als er meer dan een niveau wordt gebruikt dan moeten de weegfactor indicatoren gescheiden zijn door een puntkomma ';'.
  * Kolom _jaar_indicator_ toegevoegd waarin kan worden aangegeven welke indicator de informatie bevat over het jaartal (of voor een indicator met een andere tijdsaanduiding).
* **Slideconfiguratie**: kolommen toegevoegd en naamgeving aangepast omdat de originele naamgeving verwarring veroorzaakte
  * Kolom _weging_type_ toegevoegd waarin aangegeven welke weging moeten gebruikt (_geen_, _niveau_ of _indicator_)
  * Kolom _weegfactor_indicator_ toegevoegd die alleen gebruikt wordt als op indicator niveau wordt gewogen
  * Naam van de kolom _waarden_ is verandert naar _waarde_
  * Naam van de kolom _value_label_ is verandert naar _indicator_label_
  * Naam van de kolom _index_ is verandert naar _slidenummer_
  * Naam van de kolom _label_ is verandert naar _naam_aanduiding_
  * type _combi_ bestaat niet meer en is aangepast naar _staafgrafiek_
  * type _combi liggend_ bestaat niet meer en is aangepast naar _staafgrafiek liggend_
