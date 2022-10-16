import pandas as pd
from faker import Factory,providers
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
    fichier2=pd.read_excel("Sondage.xlsx",1)
    administre=fichier2["Administré.e"][fichier2.shape[0]-1]
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
    gen=Factory.create('fr_FR')
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
        liste_al1.append(liste_prod.pop(random.randint(0,len(liste_prod))))
        liste_al2.append(liste_prod.pop(random.randint(0,len(liste_prod))))
        liste_al3.append(liste_prod.pop(random.randint(0,len(liste_prod))))
        liste_al4.append(liste_prod.pop(random.randint(0,len(liste_prod))))
        liste_al5.append(liste_prod.pop(random.randint(0,len(liste_prod))))
        liste_al6.append(liste_prod.pop(random.randint(0,len(liste_prod))))
        liste_al7.append(liste_prod.pop(random.randint(0,len(liste_prod))))
        liste_al8.append(liste_prod.pop(random.randint(0,len(liste_prod))))
        liste_al9.append(liste_prod.pop(random.randint(0,len(liste_prod))))
        liste_al10.append(liste_prod.pop(random.randint(0,len(liste_prod))))
    df=pd.DataFrame([liste_adm,liste_lname,liste_fname,liste_bday,liste_add,liste_zcode,liste_city,liste_num,liste_al1,liste_al2,liste_al3,liste_al4,liste_al5,liste_al6,liste_al7,liste_al8,liste_al9,liste_al10])    
    df.to_excel("Sondage.xlsx",sheet_name="Feuil2",index=False,header=False,startrow=fichier2.shape[0])
sonder(10)