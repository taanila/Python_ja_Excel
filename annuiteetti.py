# Päivitetty 2025-10-19 / Aki Taanila

import pandas as pd
import numpy_financial as npf
from math import ceil
import matplotlib.pyplot as plt
from collections import OrderedDict
from datetime import date
from dateutil.relativedelta import relativedelta


def generoi_lyhennys(aloituspäivä, laina, korkokanta, maksuerä):

    '''
    Funktio generoi lyhennykset ja korot sekä palauttaa lyhennystaulukon OrderedDict-oliona.
    lyhennystaulukko-funktio kutsuu tätä funktiota ja muuntaa OrderedDict-olion dataframeksi.
    '''

    p = 1   #periodi
    ennen_lyhennystä = laina
    lyhennyksen_jälkeen = laina
    
    while lyhennyksen_jälkeen > 0: 
        korko = round(((korkokanta/12) * ennen_lyhennystä), 2)
        maksuerä = min(maksuerä, ennen_lyhennystä + korko)
        lyhennys = maksuerä - korko
        lyhennyksen_jälkeen = ennen_lyhennystä - lyhennys
        kuukausi = f'{int(aloituspäivä.year)}-{int(aloituspäivä.month)}'
        yield OrderedDict([('Periodi', p),
                           ('Kuukausi', kuukausi),
                           ('Laina ennen lyhennystä', ennen_lyhennystä),
                           ('Maksuerä', maksuerä),
                           ('Korko', korko),
                           ('Lyhennys', lyhennys),
                           ('Laina lyhennyksen jälkeen', lyhennyksen_jälkeen)])
        p += 1
        aloituspäivä += relativedelta(months=1)  #sama päivä seuraavassa kuussa
        ennen_lyhennystä = lyhennyksen_jälkeen


def lyhennystaulukko(aloituspäivä, laina, korkokanta, aika_kuukausina, maksuerä):

    '''
    lyhennystaulukko-funktio kutsuu generoi_lyhennys-funktiota, muuntaa tulokset 
    dataframeksi ja laskee lainaan liittyviä tietoja.

    aloituspäivä: lainan nostopäivä päivämäärämuodossa.
    laina: lainan suuruus.
    korkokanta: vuotuinen korkokanta desimaalimuodossa (esim. 2,24 % korkoa vastaava korkokanta 0.0245).
    aika_kuukausina: laina-aika kuukausina; jos 0, niin lasketaan laina-aika.
    maksuerä: kuukausittaisen maksuerän suuruus; jos 0, niin lasketaan maksuerä.
    '''
    ok = True
    
    # Käytetään numpy_financial paketin pmt- ja nper-funktioita maksuerän ja laina-ajan laskemiseen
    if (maksuerä == 0) and (aika_kuukausina > 0):
        try:
            maksuerä = -round(npf.pmt(korkokanta/12, aika_kuukausina, laina), 2)
        except:
            ok = False
    elif (maksuerä > 0) and (aika_kuukausina == 0):
        try:
            aika_kuukausina = ceil(npf.nper(korkokanta/12, -maksuerä, laina))
        except:
            ok = False
    else:
        ok = False

    # generoi_lyhennys-funktion tulokset dataframeksi
    if ok == True:
        taulukko = pd.DataFrame(generoi_lyhennys(aloituspäivä, laina, korkokanta, maksuerä))

        # Maksuerät, korot ja lyhennykset yhteensä koko laina-ajalta
        taulukko.loc['Yhteensä', 'Maksuerä'] = taulukko['Maksuerä'].sum()
        taulukko.loc['Yhteensä', 'Korko'] = taulukko['Korko'].sum()
        taulukko.loc['Yhteensä', 'Lyhennys'] = taulukko['Lyhennys'].sum()

        # Yleiset tiedot
        loppu_päivä = taulukko['Kuukausi'].iloc[-2]
        tiedot = pd.Series([laina, korkokanta, aika_kuukausina, maksuerä, 
                        taulukko.loc['Yhteensä', 'Korko'], loppu_päivä],
                        index = ['Lainan suuruus',  'Korkokanta', 'Aika kuukausina', 
                                 'Maksuerä', 'Korot yhteensä', 'Viimeinen maksuerä'])
    else:
        taulukko = 'Tarkista lähtötiedot!'
        tiedot = pd.Series([laina, korkokanta, aika_kuukausina, maksuerä, '?', '?'],
                          index = ['Lainan suuruus',  'Korkokanta', 'Aika kuukausina', 
                                 'Maksuerä', 'Korot yhteensä', 'Viimeinen maksuerä'])
        print('Tarkista lähtötiedot!')

    return taulukko, tiedot