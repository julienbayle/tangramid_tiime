import configparser
from typing import Final
import pandas as pd
from pandas import ExcelWriter
import numpy as np
from itertools import combinations
from datetime import datetime
import re

# Objectif :
# Lire les fichiers de l'export Tiime et en faire un fichier pivot
# pour lecture des données depuis un Excel de synthèse

# Lecture des paramètres

config = configparser.ConfigParser()
config.read('gestion.ini')

transactionsFile=config['Tiime']['TransactionsFilePath']
expenditureJusticationsFile=config['Tiime']['ExpenditureJusticationsFilePath']
expenditureJustificationsIgnoreIds = {int(x) if x.isdigit() else 0 for x in config['Tiime']['ExpenditureJustificationsIgnoreIds'].replace(" ", "").split(',')}
invoicesFile=config['Tiime']['InvoicesFilePath']

analysisFile=config['AnalysisFile']['FilePath']
analysisFileMainSheetName=config['AnalysisFile']['MainSheetName']
analysisFileTagColumnIndex=int(config['AnalysisFile']['TagsColumnIndex'])
validTags=set(config['AnalysisFile']['ValidTag'].lower().replace(" ", "").split(','))
ignoreLineWithTag=config['AnalysisFile']['IgnoreLineWithTag']
priorityTags=config['AnalysisFile']['priorityTags']

companyStartDate=datetime.strptime(config['Company']['StartDate'], '%d/%m/%Y')
people = config['Company']['People'].replace(" ", "").split(',')

# Etape 1 - Lecture du fichier Excel pour récupérer les tags valides

df = pd.read_excel(analysisFile, sheet_name=analysisFileMainSheetName)

for cell in np.asarray(df[df.columns[analysisFileTagColumnIndex]]):
  tags = sorted(tuple(filter(lambda term: len(term)>0 and term[0]=='#', str(cell).split())))
  if tags:
    validTags.update([" ".join(map(lambda term: term[1:], tags))])

print("Liste des tags valides : ", validTags)


# Function pour génération des couples de tags valides
# En pratique il peut y avoir plusieurs tags à ignoré saisis dans Tiime

def prepareTags(t):
  tags = t.lower().replace(" ", "")
  tags = tags.split('|') if '|' in tags else tags.split(',')
  return set([' '.join(a) if type(a) is tuple else a for a in tags + list(combinations(sorted(tags),2))])


# Etape 2.1 - Lecture et préparation du fichier des transactions

transactionsDF = pd.read_excel(transactionsFile)
transactionsDF = transactionsDF.rename(columns={
  "Intitulé de la transaction" : "Intitulé", 
  "Date de la transaction" : "Date"})
transactionsDF = transactionsDF.drop(['Id', 'IBAN du compte', "Date de l'opération"], axis=1)

transactionsDF.Reçus = transactionsDF.Reçus.replace(np.nan, "")
transactionsDF.Factures = transactionsDF.Factures.replace(np.nan, "")
transactionsDF.Commentaire = transactionsDF.Commentaire.replace(np.nan, "").apply(lambda t: t.lower())
transactionsDF.Tags = transactionsDF.Tags.replace(np.nan, "")

transactionsDF['Montant TVA'] = 0
transactionsDF['Montant HT'] = transactionsDF['Montant TTC']
transactionsDF['Mois'] = transactionsDF['Date'].apply(lambda t: (t.replace(day=1) if t > companyStartDate else companyStartDate).strftime("%d/%m/%Y"))
transactionsDF['Date'] = transactionsDF['Date'].apply(lambda t: t.strftime("%d/%m/%Y"))


# Etape 2.2 - Afficher une synthèse de l'avancement du pointage comptable

nbOK = transactionsDF.Commentaire.str.count('ok.*').sum()
print("Nombre de transactions OK : ", nbOK)

for name in people:
  nb = transactionsDF.Commentaire.str.count('action '+name+'.*').sum()
  print("Nombre de transactions avec action de la part de " + name + " : ", nb)
  nbOK = nbOK + nb

print("Nombre de transactions à traiter :", len(transactionsDF.index) - nbOK)

# Etape 2.3 - Appliquer automatiquement le tag #encaissement pour toutes les lignes associées à une facture

hasInvoice = transactionsDF.Factures.apply(lambda t: len(t) > 0)
transactionsDF.Tags = transactionsDF.Tags.map(lambda t: "encaissement").where(hasInvoice,  transactionsDF.Tags)


# Etape 2.4 - Appliquer le tag unique #nonaffecte si aucun tag valide, sinon, ne conserver que les tags valides

transactionsDF['FinalTag'] = transactionsDF.Tags.apply(prepareTags).apply(lambda tags: 'nonaffecte' if len(tags & validTags) == 0 else ' '.join(tags & validTags))
transactionsDF['FinalTag'] = transactionsDF.FinalTag.apply(lambda t: ' '.join(map(lambda n: '#'+n, t.split(' '))))


# Etape 2.5 - Pour éviter les lignes en doublons, ne pas conserver les transactions liées à un justificatif
#             De plus, ignorer les lignes lié à des opération d'annule et remplace (debit / remboursement)

noJustification = transactionsDF.Reçus.apply(lambda t: len(t) == 0)
doNotIgnore = transactionsDF.Tags.apply(lambda tags: ignoreLineWithTag not in tags)
filteredTransactionsDF = transactionsDF[noJustification & doNotIgnore].drop(["Reçus"], axis=1)
#print(filteredTransactionsDF.head(100))

nafCount = filteredTransactionsDF.FinalTag.str.count('nonaffecte').sum()
print("Nombre de transactions avec justificatif ou facture associée : ", len(transactionsDF.index) - len(filteredTransactionsDF.index))
print("Nombre de transactions sans justificatif ou facture et sans affectation : ", nafCount)
print("Nombre de transactions bien affectée sans justificatif : ", len(filteredTransactionsDF.index) - nafCount)


# Etape 3.1 - Lecture et préparation du fichier des justificatifs

justifDF = pd.read_excel(expenditureJusticationsFile)

justifDF = justifDF.drop(["Date d'ajout", "Source"], axis=1)
justifDF = justifDF.rename(columns={
  "Intitulé de la transaction" : "Intitulé", 
  "Date du reçu" : "Date"})

justifDF.Commentaire = justifDF.Commentaire.replace(np.nan, "").apply(lambda t: t.lower())
justifDF["Type d'opération"] = "Justificatif"
justifDF['Transactions bancaires'] = justifDF['Transactions bancaires'].replace(np.nan, "")
justifDF['Mois'] = justifDF['Date'].apply(lambda t: (t.replace(day=1) if t > companyStartDate else companyStartDate).strftime("%d/%m/%Y"))
justifDF['Date'] = justifDF['Date'].apply(lambda t: t.strftime("%d/%m/%Y"))

justifDF['FinalTag'] = justifDF.Tags.replace(np.nan, "").apply(prepareTags)
justifDF['FinalTag'] = justifDF.FinalTag.apply(lambda tags: 'nonaffecte' if len(tags & validTags) == 0 else ' '.join(tags & validTags))
justifDF['FinalTag'] = justifDF.FinalTag.apply(lambda t: ' '.join(map(lambda n: '#'+n, t.split(' '))))

justifDF['Montant TTC'] = -justifDF['Montant TTC']
justifDF['Montant TVA'] = -justifDF['Montant TVA']
justifDF['Montant HT'] = -justifDF['Montant HT']

# Etape 3.2 - On ignore dans l'export les justificatifs sans montant (exemple : justificatif de TVA à 0)
#             Car ils sont légitimement pas associés à une transaction et n'apporte rien à la gestion analytique
#             Puis on force la suppression de justificatifs dont la suppression dans Tiime ne fonctionne pas 
#             (cause : bug tiime)

notZero = justifDF['Montant TTC'].apply(lambda s: s != 0)
notIgnored = justifDF['Id'].apply(lambda id: int(id) not in expenditureJustificationsIgnoreIds)

filteredJustifDF = justifDF[notZero & notIgnored]
filteredJustifDF = filteredJustifDF.drop(["Id"], axis=1)
#print(justifDF.head(100))

# Etape 3.3 - Synthèse des justificatifs

nbOK = filteredJustifDF.Commentaire.str.count('ok.*').sum()
print("Nombre de justificatifs OK : ", nbOK)

for name in people:
  nb = filteredJustifDF.Commentaire.str.count('action '+name+'.*').sum()
  print("Nombre de justificatifs avec action de la part de " + name + " : ", nb)
  nbOK = nbOK + nb

print("Nombre de justificatifs à traiter :", len(filteredJustifDF.index) - nbOK)

nafCount = filteredJustifDF.FinalTag.str.count('nonaffecte').sum()
print("Nombre de justificatifs bien affectés : ", len(filteredJustifDF.index) - nafCount)
print("Nombre de justificatifs mal affectés : ", nafCount)

noTransactionDF = filteredJustifDF[filteredJustifDF["Transactions bancaires"].apply(lambda t: len(t) == 0)]
print("Nombre de justificatifs sans transaction liée", len(noTransactionDF.index))
print("Montant total TTC pour les justificatifs sans transaction liée", noTransactionDF["Montant TTC"].sum())

# Etape 4 - Ajout des informations de facturation

invoicesDF = pd.read_excel(invoicesFile)

invoicesDF = invoicesDF.drop(["id", "pays du client", "Transaction(s) liée(s)"], axis=1)
invoicesDF = invoicesDF.drop(list(filter(lambda t: re.match('montant TVA ([0-9\.]+)%', t), invoicesDF.columns)), axis=1)
invoicesDF = invoicesDF.rename(columns={
  "Intitulé de la transaction" : "Intitulé", 
  "Transaction(s) liée(s)" : "Transactions bancaires",
  "date de facture" : "Date",
  "montant TTC " :  "Montant TTC",
  "montant HT " : "Montant HT"})

invoicesDF["numéro de facture"] = invoicesDF["numéro de facture"].replace(np.nan, "").apply(lambda t: str(t))
invoicesDF["Intitulé"] = invoicesDF["numéro de facture"] + " / " + invoicesDF["nom du client"] + " / " + invoicesDF["intitulé"]
invoicesDF = invoicesDF.drop(["numéro de facture", "nom du client", "intitulé"], axis=1)

invoicesDF["Type d'opération"] = "Facture"
invoicesDF.Date = invoicesDF.Date.apply(lambda t: datetime.strptime(str(t), '%Y-%m-%d 00:00:00'))
invoicesDF['Mois'] = invoicesDF['Date'].apply(lambda t: (t.replace(day=1) if t > companyStartDate else companyStartDate).strftime("%d/%m/%Y"))
invoicesDF['Date'] = invoicesDF['Date'].apply(lambda t: t.strftime("%d/%m/%Y"))
invoicesDF['FinalTag'] = '#facture'

invoicesDF['Montant TVA'] = invoicesDF['Montant TTC'] - invoicesDF['Montant HT']

sendInvoice = invoicesDF['statut de la facture'].apply(lambda t: t != 'Brouillon')
filteredInvoiceDF = invoicesDF[sendInvoice].drop(["statut de la facture"], axis=1)
#print(filteredInvoiceDF.head())

print("Nombre de factures envoyées", len(filteredInvoiceDF.index))
print("Montant total des factures HT", filteredInvoiceDF['Montant HT'].sum())

# Etape 5 - Fusion des informations et gestion des tags multiples

dataDF = pd.concat([filteredJustifDF, filteredTransactionsDF, filteredInvoiceDF])

def electTag(t):
  tags = t.split(' ')
  for tag in tags:
    for ptag in priorityTags:
      if ptag in tag:
        return tag
  return tags[0]

multipleTags = dataDF.FinalTag.apply(lambda t: len(t.split(' ')))
nbMultiple = multipleTags[multipleTags > 1].count()
if nbMultiple > 0:
  print("Nombre de lignes avec plusieurs tags : ", nbMultiple)
  print("Sélection automatique d'un tag, résultat :")
  
  updateDF = dataDF[multipleTags > 1][["Date", "Intitulé", "FinalTag"]]

  dataDF.FinalTag = np.where(multipleTags > 1, dataDF.FinalTag.apply(lambda t: electTag(t)), dataDF.FinalTag)
  updateDF["SelectedTag"] = dataDF.FinalTag[multipleTags > 1]
  print(updateDF.rename(columns={"FinalTag":"Tags"}))

with ExcelWriter("data.xlsx") as writer:
  dataDF.to_excel(writer)

input("Export terminé sans erreur, taper entrer pour terminer...")