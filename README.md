# Pythonilla lasketut tulokset Exceliin

*Voit ladata esimerkit omalle koneelle: Napsauta **Code** (vihreä pudotusvalikko-painike) ja valitse **Download ZIP**. Pura ZIP-paketti itsellesi sopivaan hakemistoon.*

Excelissä on helppokäyttöinen käyttöliittymä tiedon järjestelyyn, esittämiseen ja laskentaan. Python tarjoaa Exceliä monipuolisemmat laskentamahdollisuudet erityisesti data-analytiikkaan. Onneksi Pythonin ja Excelin yhteiskäyttö on nykyään mahdollista. Voit hyödyntää Pythonin monipuolisia laskentamahdollisuuksia ja kirjoittaa tulokset Exceliin.

On olemassa erilaisia tapoja siirtää Pythonilla laskettuja tuloksia Exceliin tai Excel-tiedostoon. On tärkeää huomata, että tiedon kirjoittaminen avoinna olevaan Exceliin on teknisesti täysin eri asia kuin tiedon kirjoittaminen Excel-tiedostoon. 

Suoraan avoimeen Exceliin kirjoittaminen mahdollistaa vuorovaikutteisen Jupyter notebookin ja Excelin käytön. Suoraan avoimeen Exceliin kirjoitettaessa täytyy huomioida käyttöjärjestelmien (Windows, MacOS) erot. Avoimeen Exceliin kirjoitettaessa **xlwings**-paketti häivyttää suurimman osan käyttöjärjestelmien eroista. Käyttäjän vaivannäköä tarvitaan kuitenkin joissain harvemmin tarvittavissa muotoilutoiminnoissa. **xlwings**-paketti on valmiiksi asennettu Anacondaan. Minicondaan voit asentaa sen komennolla `conda install xlwings`. Jos käytätä pip-paketinhallintaa, niin asennus hoituu komennolla `pip install xlwings`.

Excel-tiedostot ovat samanlaisia riippumatta käyttöjärjestelmästä ja tiedostojen välityksellä tiedot voidaan avata myös muihin taulukkolaskentaohjelmiin. Suoraan tiedostoon kirjoitettaessa Excelin ei tarvitse olla edes asennettuna koneelle. Käytän **xlsxwriter**-moduulia tiedon kirjoittamiseen suoraan Excel-tiedostoon. **xlsxwriter**in asennus hoituu komennolla `conda install xlsxwriter` tai `pip install xlsxwriter`. **Pandas**-paketin kautta käytettynä **xlsxwriter** kirjoittaa kerralla kokonaisen dataframen tiedostoon.

Testaan esimerkkien avulla neljää erilaista tapaa käyttää Pythonilla laskettuja tuloksia Excelissä.. Jokaisessa esimerkissä käytän laskentana annuiteettilainan lyhennystaulukon laskemista ja eri lähtötiedoilla varustettujen lainojen vertailua. Laskennan hoidan kahden funktion avulla, jotka olen määritellyt **annuiteetti.py**-moduulissa. Esimerkkien koodi toimii, jos **annuiteetti.py** on samassa kansiossa jupyter-muistion kanssa. Excel-esimerkissä **laina_xlwings_lite.xlsx** laskentaan käytettävät Python-funktiot on tallennettu Excel-tiedostoon. Varsinaisena aiheena on tulosten kirjoittaminen Exceliin, joten laskentaan käyttämäni funktiot ovat sivuosassa. Kannattaa kuitenkin tutustua funktioihin **annuiteetti.py**-moduulissa. Lienee myös hyödyllistä tutustua ennen Excel-esimerkkejä muistioon **laina.ipynb**, jossa suoritetaan esimerkeissä toistuvat laskelmat ilman tulosten siirtämistä Exceliin.

## xlwings-paketin view-funktio

**xlwings**-paketin **view**-funktio (https://docs.xlwings.org/en/latest/jupyternotebooks.html) on nopein ja helpoin tapa kirjoittaa tietoa suoraan Exceliin. Funktio tekee samalla kertaa monta asiaa: 
- avaa koneelle asennetun Excelin
- luo uuden työkirjan
- kirjoittaa **pandas**-tietorakenteen (**dataframe** tai **series**) työkirjaan.

Jos **xlwings**-paketti on tuotu **xw**-nimisenä, niin **view**-funktio siirtää esimerkiksi **taulukko**-nimisen dataframen Exceliin komennolla `xw.view(taulukko)`. Opi lisää esimerkkimuistiosta **laina_view.ipynb**.

## xlwings-paketti

Vaativampaan tiedon kirjoittamiseen ja muotoiluun voit käyttää muita **xlwings**-paketin toimintoja (https://docs.xlwings.org/en/latest/). Tällöin sinun pitää luoda **xlwings**illä piilotettu Excel-instanssi `app=xw.App(visible=False)` ja Excel-työkirja `wb=app.books[0]`. Ilman piilotusta hätäinen käyttäjä ehtii kajota Excel-tiedostoon tiedon kirjoituksen ja muotoilun aikana, josta yleensä seuraa ongelmia. Tietoa voi kirjoittaa **value**-ominaisuuden avulla, esimerkiksi `ws1.range(1, sarake).value = 'Annuiteettilainan lyhennystaulukko'. Opi lisää esimerkkimuistiosta **laina_xlswings.ipynb**.

**xlwings**-paketin toiminnoilla et voi kirjoittaa suoraan Excel-tiedostoon. Voit kuitenkin halutessasi tallentaa **xlwings**illä luodun työkirjan **save**-funktiolla, joka vastaa tiedoston tallentamista Excelin **Save As** -toiminnolla. Jos et määritä hakemistopolkua ja tiedoston nimeä, niin työkirja tallennetaan oletusnimellä sen hetkiseen oletushakemistoon. Jos hakemistossa on entuudestaan saman niminen tiedosto, niin se ylikirjoitetaan.

## xlwings_lite

**xlwings_lite** on uusin (ensimmäinen versio julkaistu 21.3.2025) tapa Pythonin ja Excelin yhteiskäyttöön. **xlwings_lite** tuo Pythonin Exceliin. Voit kirjoittaa koodia ja käyttää **xlwings**-paketin toimintoja Excelin sisällä. Samalla saat käyttöösi kaikki Pythonin tarjoamat monipuoliset laskentamahdollisuudet koneoppimisen mallit mukaan lukien. Et tarvitse omalle koneelle edes Python-asennusta. Useimmissa tapauksissa voit käyttää **xlwings_lite**ä myös **VBA**-kielisten makrojen sijasta Excelin automatisoimiseen. 

Voit asentaa **xlwings_lite**n Excelin **Add-ins** (Apuohjelmat) -toiminnolla. Asennusohje: https://lite.xlwings.org/installation. Asennuksen jälkeen Excelin työkalunauhassa on **xlwings_Lite**-painike, josta aukeaa **xlwings_Lite**-paneeli. Paneelissa on kaksi välilehteä: **main.py** ja **requirements.txt**. Asennetut paketit/kirjastot löydät **requirements.txt**-välilehdeltä. Voit lisätä paketin nimen listan viimeiseksi. Jos lisääminen ei aiheuta virheilmoitusta, niin voit käyttää kyseistä pakettia koodissasi. 

Opi lisää tutustumalla Excel-esimerkkiin **laina_xlswings_lite.xlsx**. Huomaa, että **xlwings_lite**ä käytettäessä kaikki tarpeellinen on tallennettu kätevästi yhteen Excel-tiedostoon. **xlwings_Lite** välilehdelle olen lisännyt paketin **numpy-financial**.

## xlsxwriter kirjoittaa suoraan Excel-tiedostoon

Excel-tiedoston luomiseen ja kirjoittamiseen tarkoitettu **xlsxwriter**-moduuli kirjoittaa 100 % yhteensopivia Excel-tiedostoja, joissa on mahdollista käyttää monipuolisesti Excelin muotoiluja. Moduuli on hyvin dokumentoitu osoitteessa https://xlsxwriter.readthedocs.io/.

**Pandas**-kirjaston **to_excel**-funktion osaa käyttää kirjoittimena **xlsxwriter**ia jopa kokonaisen dataframen kirjoittamiseen kerralla tiedostoon. Dataframen otsikkorivin ja index-sarakkeen muotoilu oletusmuotoilusta poikkeavalla tavalla vaatii ylimääräistä vaivannäköä. Tätä valaisevat esimerkit löytyvät osoitteesta https://xlsxwriter.readthedocs.io/pandas_examples.html. 

Suoraan tiedostoon kirjoittaminen sopii tilanteisiin, joissa ei tarvita reaaliaikaista interaktiivisuutta Pythonin ja Excelin välillä. Tällaisia ovat esimerkiksi monet raportointitehtävät.

## Microsoftin Python in Excel

Microsoftilla on oma versio Pythonin ja Excelin yhteiskäyttöön. Tämä poikkeaa **xlwings**- ja **xlsxwriter**-paketeista ainakin seuraavissa kohdissa:
- **Python in Excel** ei ainakaan tätä kirjoitettaessa ole tarjolla kaikissa Office-paketin versioissa.
- Python-koodit tallennetaan yksittäisiin Excelin soluihin. Jos koodia on paljon ja se hajaantuu ympäri työkirjaa yksittäisiin soluihin, niin kokonaisuuden hallinta voi olla vaikeaa.
- Koodi lähetetään aina suoritettavaksi Microsoftin pilvipalveluun. Samalla lähetetään pilvipalveluun myös data, jota koodi käsittelee.

Lisätietoa: https://support.microsoft.com/fi-fi/office/johdanto-pythoniin-exceliss%C3%A4-55643c2e-ff56-4168-b1ce-9428c8308545














