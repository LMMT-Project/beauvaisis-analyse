import pandas as pd
from faker import Factory
import random
#lecture du fichier 
fichier =  pd.read_excel("Aliments.xlsx")

#dictionnaire -> clé : code produit, valeur : catégorie(s) 
dic_cat = {}

#parcours du fichier ligne par ligne
for i in range(fichier.shape[0]):
    liste_cat=[]
    temp = fichier.alim_nom_fr[i].split()
    for j in temp:
        if j.strip()=="bio":
            liste_cat.append("bio")
        if j.strip()=="vegan":
            liste_cat.append("vegan")
        temp=fichier.alim_ssssgrp_nom_fr[i].split()
    for j in temp:
        if j.strip()=="casher":
                liste_cat.append("casher")
                liste_cat.append("halal")
        elif j.strip()=="halal":
                liste_cat.append("halal")
    if len(liste_cat)!=0:
        dic_cat[fichier.alim_code[i]]=dic_cat
print(dic_cat)

def sonder(n):
    fichier2=pd.read_excel("Sondage.xlsx")
    liste_adm=[]
    liste_lname=[]
    liste_fname=[]
    liste_bday=[]
    liste_add=[]
    liste_zcode=[]
    liste_city=[]
    liste_num=[]
    liste_al1=[]
    liste_al2=[]
    liste_al3=[]
    liste_al4=[]
    liste_al5=[]
    liste_al6=[]
    liste_al7=[]
    liste_al8=[]
    liste_al9=[]
    liste_al10=[]
    for i in range(fichier2.shape[0]):
        liste_adm.append(fichier2["Administré.e"][i])
        liste_lname.append(fichier2["Nom"][i])
        liste_fname.append(fichier2["Prénom"][i])
        liste_bday.append(fichier2["Naissance"][i])
        liste_add.append(fichier2["Adresse"][i])
        liste_zcode.append(fichier2["Code Postal"][i])
        liste_city.append(fichier2["Ville"][i])
        liste_num.append(fichier2["Tel"][i])
        liste_al1.append(fichier2["Aliment1"][i])
        liste_al2.append(fichier2["Aliment2"][i])
        liste_al3.append(fichier2["Aliment3"][i])
        liste_al4.append(fichier2["Aliment4"][i])
        liste_al5.append(fichier2["Aliment5"][i])
        liste_al6.append(fichier2["Aliment6"][i])
        liste_al7.append(fichier2["Aliment7"][i])
        liste_al8.append(fichier2["Aliment8"][i])
        liste_al9.append(fichier2["Aliment9"][i])
        liste_al10.append(fichier2["Aliment10"][i])
    gen=Factory.create('fr_FR')
    administre=liste_adm[-1]
    for i in range(n):
        administre+=1
        liste_prod=fichier.alim_code.tolist()
        adresse=gen.address().split("\n")
        liste_adm.append(administre)
        liste_lname.append(gen.last_name())
        liste_fname.append(gen.first_name())
        liste_bday.append(gen.date_of_birth(None,12).strftime("%Y/%m/%d"))
        liste_add.append(adresse[0])
        liste_zcode.append(adresse[1].split()[0])
        liste_city.append(adresse[1].split()[1])
        liste_num.append(gen.phone_number())
        liste_al1.append(liste_prod.pop(random.randint(0,len(liste_prod)-1)))
        liste_al2.append(liste_prod.pop(random.randint(0,len(liste_prod)-1)))
        liste_al3.append(liste_prod.pop(random.randint(0,len(liste_prod)-1)))
        liste_al4.append(liste_prod.pop(random.randint(0,len(liste_prod)-1)))
        liste_al5.append(liste_prod.pop(random.randint(0,len(liste_prod)-1)))
        liste_al6.append(liste_prod.pop(random.randint(0,len(liste_prod)-1)))
        liste_al7.append(liste_prod.pop(random.randint(0,len(liste_prod)-1)))
        liste_al8.append(liste_prod.pop(random.randint(0,len(liste_prod)-1)))
        liste_al9.append(liste_prod.pop(random.randint(0,len(liste_prod)-1)))
        liste_al10.append(liste_prod.pop(random.randint(0,len(liste_prod)-1)))
    s1=pd.Series(liste_adm,name="Administré.e")
    s2=pd.Series(liste_lname,name="Nom")
    s3=pd.Series(liste_fname,name="Prénom")
    s4=pd.Series(liste_bday,name="Naissance")
    s5=pd.Series(liste_add,name="Adresse")
    s6=pd.Series(liste_zcode,name="Code Postal")
    s7=pd.Series(liste_city,name="Ville")
    s8=pd.Series(liste_num,name="Tel")
    s9=pd.Series(liste_al1,name="Aliment1")
    s10=pd.Series(liste_al2,name="Aliment2")
    s11=pd.Series(liste_al3,name="Aliment3")
    s12=pd.Series(liste_al4,name="Aliment4")
    s13=pd.Series(liste_al5,name="Aliment5")
    s14=pd.Series(liste_al6,name="Aliment6")
    s15=pd.Series(liste_al7,name="Aliment7")
    s16=pd.Series(liste_al8,name="Aliment8")
    s17=pd.Series(liste_al9,name="Aliment9")
    s18=pd.Series(liste_al10,name="Aliment10")
    df=pd.concat([s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11,s12,s13,s14,s15,s16,s17,s18],axis=1)  
    df.to_excel("Sondage.xlsx",sheet_name="Résultat Sondage",index=False)
sonder(10541)