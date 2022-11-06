import pandas as pd
from faker import Factory
import random
import matplotlib.pyplot as plt
excel_aliments =  pd.read_excel("Aliments.xlsx")

#dictionnaire -> clé : code produit, valeur : catégorie(s) 
categories = {}
cat_prod={}
calories={}
calo_prod={}
#parcours du excel_aliments ligne par ligne
for i in range(excel_aliments.shape[0]):
    liste_categorie = []
    cat_prod[excel_aliments.alim_code[i]] = excel_aliments.alim_grp_nom_fr[i]
    calo_prod[excel_aliments.alim_code[i]] = str(excel_aliments["Energie, Règlement UE N° 1169/2011 (kcal/100 g)"][i]).replace(",", ".").strip()
    mots_alim_nom_fr = excel_aliments.alim_nom_fr[i].split()
    for mot in mots_alim_nom_fr:
        if mot.strip() in ["bio", "vegan"]:
            liste_categorie.append(mot.strip())
    
    mots_ssssgrp_nom_fr = excel_aliments.alim_ssssgrp_nom_fr[i].split()
    for mot in mots_ssssgrp_nom_fr:
        if mot.strip() in ["casher", "halal"]:
            liste_categorie.append(mot.strip())

    if len(liste_categorie) != 0:
        categories[excel_aliments.alim_code[i]] = liste_categorie



def sonder(n):
    excel_sondage=pd.read_excel("Sondage.xlsx")

    liste_adm = []
    liste_lname = []
    liste_fname = []
    liste_bday = []
    liste_add = []
    liste_zcode = []
    liste_city = []
    liste_num = []
    liste_al1 = []
    liste_al2 = []
    liste_al3 = []
    liste_al4 = []
    liste_al5 = []
    liste_al6 = []
    liste_al7 = []
    liste_al8 = []
    liste_al9 = []
    liste_al10 = []
    #copie du fichier
    for i in range(excel_sondage.shape[0]):
        liste_adm.append(excel_sondage.get("Administré.e")[i])
        liste_lname.append(excel_sondage["Nom"][i])
        liste_fname.append(excel_sondage["Prénom"][i])
        liste_bday.append(str(excel_sondage["Naissance"][i]))
        liste_add.append(excel_sondage["Adresse"][i])
        liste_zcode.append(excel_sondage["Code Postal"][i])
        liste_city.append(excel_sondage["Ville"][i])
        liste_num.append(excel_sondage["Tel"][i])
        liste_al1.append(excel_sondage["Aliment1"][i])
        liste_al2.append(excel_sondage["Aliment2"][i])
        liste_al3.append(excel_sondage["Aliment3"][i])
        liste_al4.append(excel_sondage["Aliment4"][i])
        liste_al5.append(excel_sondage["Aliment5"][i])
        liste_al6.append(excel_sondage["Aliment6"][i])
        liste_al7.append(excel_sondage["Aliment7"][i])
        liste_al8.append(excel_sondage["Aliment8"][i])
        liste_al9.append(excel_sondage["Aliment9"][i])
        liste_al10.append(excel_sondage["Aliment10"][i])

    gen = Factory.create('fr_FR')
    administre = liste_adm[-1]
    #ajout de nouvelle ligne
    for i in range(n):
        administre += 1
        liste_prod = excel_aliments.alim_code.tolist()
        adresse = gen.address().split("\n")
        liste_adm.append(administre)
        liste_lname.append(gen.last_name())
        liste_fname.append(gen.first_name())
        liste_bday.append(gen.date_of_birth(None, 12).strftime("%d/%m/%Y"))
        liste_add.append(adresse[0])
        liste_zcode.append(adresse[1].split()[0])
        liste_city.append(adresse[1].split()[1])
        liste_num.append(gen.phone_number())
        liste_al1.append(liste_prod.pop(random.randint(0, len(liste_prod)-1)))
        liste_al2.append(liste_prod.pop(random.randint(0, len(liste_prod)-1)))
        liste_al3.append(liste_prod.pop(random.randint(0, len(liste_prod)-1)))
        liste_al4.append(liste_prod.pop(random.randint(0, len(liste_prod)-1)))
        liste_al5.append(liste_prod.pop(random.randint(0, len(liste_prod)-1)))
        liste_al6.append(liste_prod.pop(random.randint(0, len(liste_prod)-1)))
        liste_al7.append(liste_prod.pop(random.randint(0, len(liste_prod)-1)))
        liste_al8.append(liste_prod.pop(random.randint(0, len(liste_prod)-1)))
        liste_al9.append(liste_prod.pop(random.randint(0, len(liste_prod)-1)))
        liste_al10.append(liste_prod.pop(random.randint(0 ,len(liste_prod)-1)))
    #mise en forme des résultats sous forme de série 
    s1=pd.Series(liste_adm, name="Administré.e")
    s2=pd.Series(liste_lname, name="Nom")
    s3=pd.Series(liste_fname, name="Prénom")
    s4=pd.Series(liste_bday, name="Naissance",dtype=str)
    s5=pd.Series(liste_add, name="Adresse")
    s6=pd.Series(liste_zcode, name="Code Postal")
    s7=pd.Series(liste_city, name="Ville")
    s8=pd.Series(liste_num, name="Tel")
    s9=pd.Series(liste_al1, name="Aliment1")
    s10=pd.Series(liste_al2, name="Aliment2")
    s11=pd.Series(liste_al3, name="Aliment3")
    s12=pd.Series(liste_al4, name="Aliment4")
    s13=pd.Series(liste_al5, name="Aliment5")
    s14=pd.Series(liste_al6, name="Aliment6")
    s15=pd.Series(liste_al7, name="Aliment7")
    s16=pd.Series(liste_al8, name="Aliment8")
    s17=pd.Series(liste_al9, name="Aliment9")
    s18=pd.Series(liste_al10, name="Aliment10")
    df=pd.concat([s1, s2, s3, s4, s5, s6, s7, s8, s9, s10 , s11, s12, s13, s14, s15, s16, s17, s18], axis=1)  
    df.to_excel("Sondage.xlsx", sheet_name="Résultat Sondage", index=False)

    
#sonder(10000)
def cherche_categ(codeprod):
    return cat_prod[codeprod]

def regimecateg():
    excel_sondage=pd.read_excel("Sondage.xlsx")
    resultat_regime=[]
    resultat_categ={}
    for personne in range(excel_sondage.shape[0]):
        regime_personne={}
        regime_personne["bio"]=0
        regime_personne["vegan"]=0
        regime_personne["halal"]=0
        regime_personne["casher"]=0
        for aliment in range(1,11):
            nom_colonne = "Aliment"+str(aliment)
            if excel_sondage[nom_colonne][personne] in categories.keys():
                liste_categ = categories.get(excel_sondage[nom_colonne][personne])
                for categorie in liste_categ:
                    regime_personne[categorie]+=1
            categ_produit = cherche_categ(excel_sondage[nom_colonne][personne])
            if categ_produit not in resultat_categ.keys():
                resultat_categ[categ_produit] = 1
            else:
                resultat_categ[categ_produit] += 1
        maxi=0
        regime_privilegie="Aucun"
        for regime in regime_personne.keys():
            if regime_personne[regime]>maxi:
                regime_privilegie = regime
                maxi = regime_personne[regime]
        resultat_regime.append(regime_privilegie)
    return resultat_regime, resultat_categ

regime_pers, prod_categ = regimecateg()
print(regime_pers)
print(prod_categ)

def moyenneCaloriesAliments():
    nbAlimValable = 0
    totalCalories = 0
    for aliment in calo_prod.keys():
        if calo_prod[aliment] != "-" and calo_prod[aliment] != "" and calo_prod[aliment] != "nan":
            nbAlimValable += 1
            totalCalories += float(calo_prod[aliment])
            if(totalCalories != totalCalories):
                print("cpt")
    return float(totalCalories/nbAlimValable)

def score_sante():
    excel_sondage = pd.read_excel("Sondage.xlsx")
    score_sante=[]
    moyenneCalories = moyenneCaloriesAliments()

    for personne in range(excel_sondage.shape[0]):
        score = 0
        currentCalorie = 0
        nbAlimValable = 0
        for aliment in range(1,11):
            nom_colonne = "Aliment"+str(aliment)
            codeAlim = excel_sondage[nom_colonne][personne]
            if calo_prod[codeAlim] != "-" and calo_prod[codeAlim] != "" and calo_prod[codeAlim] != "nan":
                currentCalorie += float(calo_prod[codeAlim])
                nbAlimValable += 1
        if nbAlimValable != 0: 
            moyenneRef = moyenneCalories * nbAlimValable
            
            score = currentCalorie/moyenneRef * 100
            if score >100:
                score=100-(score-100)
            if score < 0:
                score=0
            score_sante.append(round(score,2))
    return score_sante
sante_pers = score_sante()
print(sante_pers)
