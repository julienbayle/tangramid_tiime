import pandas as pd
from pandas import ExcelWriter
import numpy as np
from itertools import combinations
import re

# Objectif :
# Lire les fichiers de l'export Tiime et en faire un fichier pivot
# pour lecture des données depuis un Excel de synthèse

# Paramètres
transactionsFile='tiime/transactions_bancaire.xlsx'
expenditureJusticationFile='tiime/recus.xlsx'
analysisFile='Tableau de suivi.xlsx'
analysisSheetIndex=1
projectTagColumnIndex=2
people = ['manuela', 'vanessa']

# Etape 1 - Lecture du fichier Excel pour récupérer les tags reconnaissables

validTags=set()
df = pd.read_excel(analysisFile, sheet_name=analysisSheetIndex)

for cell in np.asarray(df[df.columns[projectTagColumnIndex]]):
  tags = sorted(tuple(filter(lambda term: len(term)>0 and term[0]=='#', str(cell).split())))
  if tags:
    validTags.update([" ".join(map(lambda term: term[1:], tags))])

print("Liste des tags valides : ", validTags)


# Etape 2.1 - Lecture et préparation du fichier des transactions

transactionsDF = pd.read_excel(transactionsFile)
transactionsDF = transactionsDF.rename(columns={ "Intitulé de la transaction" : "Intitulé" })
transactionsDF.Commentaire = transactionsDF.Commentaire.replace(np.nan, "").apply(lambda t: t.lower())
transactionsDF['FinalTag'] = transactionsDF.Tags.replace(np.nan, "").apply(lambda t: set([' '.join(a) if type(a) is tuple else a for a in t.lower().replace(" ", "").split(',') + list(combinations(sorted(t.lower().replace(" ", "").split(',')),2))]) )
transactionsDF.Reçus = transactionsDF.Reçus.replace(np.nan, "")
transactionsDF.Factures = transactionsDF.Factures.replace(np.nan, "")
transactionsDF = transactionsDF.drop(['Id', 'IBAN du compte', "Date de l'opération"], axis=1)
transactionsDF['Montant TVA'] = 0
transactionsDF['Montant HT'] = transactionsDF['Montant TTC']
transactionsDF['Mois'] = transactionsDF['Date de la transaction'].apply(lambda t: t.replace(day=1).strftime("%d/%m/%Y"))
transactionsDF['Date de la transaction'] = transactionsDF['Date de la transaction'].apply(lambda t: t.strftime("%d/%m/%Y"))


# Etape 2.2 - Afficher une synthèse de l'avancement du pointage comptable

nbOK = transactionsDF.Commentaire.str.count('ok.*').sum()
print("Nombre de transactions OK : ", nbOK)

for name in people:
  nb = transactionsDF.Commentaire.str.count('action '+name+'.*').sum()
  print("Nombre de transactions avec action de la part de " + name + " : ", nb)
  nbOK = nbOK + nb

print("Nombre de transactions à traiter :", len(transactionsDF.index) - nbOK)


# Etape 2.3 - Transaction sans justificatif, sans facture => Appliquer le tag unique #nonaffecte et aucun tag valide, sinon, ne conserver que les tags valides

noJustification = transactionsDF.Reçus.apply(lambda t: len(t) == 0)
noInvoice = transactionsDF.Factures.apply(lambda t: len(t) == 0)
transactionsDF['FinalTag'] = transactionsDF.FinalTag.apply(lambda tags: 'nonaffecte' if len(tags & validTags) == 0 else ' '.join(tags & validTags))
transactionsDF['FinalTag'] = transactionsDF.FinalTag.apply(lambda t: ' '.join(map(lambda n: '#'+n, t.split(' '))))
filteredTransactionsDF = transactionsDF[noJustification].drop(["Reçus", "Factures"], axis=1)

print(filteredTransactionsDF.head(100))

nafCount = filteredTransactionsDF.FinalTag.str.count('nonaffecte').sum()
print("Nombre de transactions avec justificatif ou facture associée : ", len(transactionsDF.index) - len(filteredTransactionsDF.index))
print("Nombre de transactions sans justificatif ou facture et sans affectation : ", nafCount)
print("Nombre de transactions bien affectée sans justificatif : ", len(filteredTransactionsDF.index) - nafCount)


# Etape 3.1 - Lecture et préparation du fichier des justificatifs

justifDF = pd.read_excel(expenditureJusticationFile)
justifDF.Commentaire = justifDF.Commentaire.replace(np.nan, "").apply(lambda t: t.lower())
justifDF["Type d'opération"] = "Justificatif"
justifDF['Transactions bancaires'] = justifDF['Transactions bancaires'].replace(np.nan, "")
justifDF = justifDF.drop(["Id", "Date d'ajout", "Source"], axis=1)
justifDF['Mois'] = justifDF['Date du reçu'].apply(lambda t: t.replace(day=1).strftime("%d/%m/%Y"))
justifDF['Date du reçu'] = justifDF['Date du reçu'].apply(lambda t: t.strftime("%d/%m/%Y"))
justifDF['FinalTag'] = justifDF.Tags.replace(np.nan, "").apply(lambda t: set([' '.join(a) if type(a) is tuple else a for a in t.lower().replace(" ", "").split('|') + list(combinations(sorted(t.lower().replace(" ", "").split('|')),2))]) )
justifDF['FinalTag'] = justifDF.FinalTag.apply(lambda tags: 'nonaffecte' if len(tags & validTags) == 0 else ' '.join(tags & validTags))
justifDF['FinalTag'] = justifDF.FinalTag.apply(lambda t: ' '.join(map(lambda n: '#'+n, t.split(' '))))
print(justifDF.head(100))

nbOK = justifDF.Commentaire.str.count('ok.*').sum()
print("Nombre de justificatifs OK : ", nbOK)

for name in people:
  nb = justifDF.Commentaire.str.count('action '+name+'.*').sum()
  print("Nombre de justificatifs avec action de la part de " + name + " : ", nb)
  nbOK = nbOK + nb

print("Nombre de justificatifs à traiter :", len(justifDF.index) - nbOK)

nafCount = justifDF.FinalTag.str.count('nonaffecte').sum()
print("Nombre de justificatifs bien affectés : ", len(justifDF.index) - nafCount)
print("Nombre de justificatifs mal affectés : ", nafCount)

noTransactionDF = justifDF[justifDF["Transactions bancaires"].apply(lambda t: len(t) == 0)]
print("Nombre de justificatifs sans transaction liée", len(noTransactionDF.index))
print("Montant total TTC pour les justificatifs sans transaction liée", noTransactionDF["Montant TTC"].sum())

with ExcelWriter("data.xlsx") as writer:
  pd.concat([justifDF, filteredTransactionsDF]).to_excel(writer)

input("Taper entrer pour terminer...")