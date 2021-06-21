library("readxl")
library("writexl")
ABCD=''

ADDITIONDATAFRAMES_SIMPLE=function(data) {
  x=data.frame(data,Pt=1:length(data[1]),Ecart=1:length(data[1]),Ecart_Absolu=1:length(data[1]))
}
Calcul_CFE_SIMPLE=function(dataset,x) {
  return(sum(dataset[,4])-x)
}

Calcul_CFE_DOUBLE=function(dataset,x) {
  return(sum(dataset[,6])-x)
}

Calcul_CFE_TRIPLE=function(dataset,x) {
  return(sum(dataset[,8])-x)
}

Calcul_MAD_SIMPLE=function(dataset,x) {
  return(((mean(dataset[,5])*16)-x)/12)
}

Calcul_MAD_DOUBLE=function(dataset,x) {
  return(((mean(dataset[,7])*16)-x)/12)
}

Calcul_MAD_TRIPLE=function(dataset,x) {
  return(((mean(dataset[,9])*16)-x)/12)
}

Calcul_MSE_SIMPLE=function(dataset,x) {
  return(((mean((dataset[,2][5:16]-dataset[,3][5:16])**2))))
}

Calcul_MSE_DOUBLE=function(dataset,x) {
  return(((mean((dataset[,2][5:16]-dataset[,5][5:16])**2))))
}

Calcul_MSE_TRIPLE=function(dataset,x) {
  return(((mean((dataset[,2][5:16]-dataset[,7][5:16])**2))))
}
Calcul_Pt_Init_SIMPLE=function(dataset,x) {
  dataset[x,3]=mean(unlist(dataset[2],use.names = FALSE)[1:x])
  return(dataset)
}

Calcul_Pt_SIMPLE=function(dataset,x,alpha) {
  for (i in 1:(length(unlist(dataset[1],use.names = FALSE))-x)) {
    dataset[i+x,3]=alpha*dataset[i+x-1,2]+(1-alpha)*dataset[i+x-1,3]
  }
  return(dataset)
}

Calcul_Ecart_EcartAbsolu_SIMPLE=function(dataset,x) {
  for (i in 1:(length(unlist(dataset[1],use.names = FALSE))-x)) {
    dataset[i+x,4]=dataset[i+x,2]-dataset[i+x,3]
    dataset[i+x,5]=abs(dataset[i+x,4])
  }
  return(dataset)
}

Lissage_Simple=function(chemin,Debut,alpha) {
  x=read_excel(chemin)
  x=ADDITIONDATAFRAMES_SIMPLE(x)
  LISSAGESIMPLE=Calcul_Ecart_EcartAbsolu_SIMPLE(Calcul_Pt_SIMPLE(Calcul_Pt_Init_SIMPLE(x,Debut),Debut,alpha),Debut)
  write_xlsx(y,"C:\\Users\\msi\\Desktop\\Résultat_Lissage_Simple.xlsx")
  return(LISSAGESIMPLE)
}

Visualisation_Graphique_Lissage_Simple=function(x) {
  A=unlist(x[2],use.names = FALSE)
  B=unlist(x[3],use.names = FALSE)
  png(file = "Graphe du lissage simple.jpg")
  plot(A,type = "o",col = "red", xlab = "Période", ylab = "Quantité", main = "Prévision de la demande, Méthode du lissage simple")
  lines(B, type = "o", col = "blue")
  dev.interactive()
  dev.off()
}
require(tcltk)
Calcul_CFE_SIMPLE=function(dataset,x) {
  return(sum(dataset[,4])-x)
}
Fenetre_Simple=function() {
  
  DebutVariable=tclVar("")
  AlphaVariable=tclVar("")
  tt=tktoplevel()
  tkwm.title(tt,"Prévision de la demande - Lissage simple")
  
  Debut.entry=tkentry(tt,textvariable=DebutVariable)
  Alpha.entry=tkentry(tt,textvariable=AlphaVariable)
  
  Lissage_Simplee=function() {
    p=Lissage_Simple(chemin,as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)))
    tkmessageBox(message=paste("Merci de trouver le resultat dans le nouveau excel crée dans le bureau sous le nom de Résultat_Lissage_Simple.xlsx, veuillez aussi créer une copie car le document va se faire écraser dans le cas d'un nouveau usage de l'application"))
    return(p)
  }
  LissageSimple.but=tkbutton(tt, text="Start", command=Lissage_Simplee)
  Calcul_CFE_SIMPLEE=function() {
    r=Calcul_MAD_SIMPLE(Lissage_Simple(chemin,as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable))),as.double(tclvalue(DebutVariable)))
    g=Calcul_CFE_SIMPLE(Lissage_Simple(chemin,as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable))),as.double(tclvalue(DebutVariable)))
    h=Calcul_MSE_SIMPLE(Lissage_Simple(chemin,as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable))),as.double(tclvalue(DebutVariable)))
    tkmessageBox(message=paste("Mean Absolute Deviation =",r))
    tkmessageBox(message=paste("Cumulative Forecast Error =",g))
    tkmessageBox(message=paste("Mean Square Error =",h))
  }
  
  Mesure_Precision.but=tkbutton(tt, text="Afficher les indicateurs mesurants la précision de la prévision", command=Calcul_CFE_SIMPLEE)
  
  
  Visu_Graphique_Simple=function() {
    Visualisation_Graphique_Lissage_Simple(Lissage_Simple(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable))))
    tkmessageBox(message=paste("Merci de trouver le graphe dans le dossier 'Documents', veuillez également créer une copie car l'image va se faire écraser dans le cas d'un nouveau usage du script"))
  }
  VisuLissSimplee.but=tkbutton(tt, text="Visualiser Les graphes Demandes/Prévisions", command=Visu_Graphique_Simple)
  
  reset=function(){
    tclvalue(DebutVariable)=4
    tclvalue(AlphaVariable)=0.3
  }
  tkgrid(tklabel(tt,text="Merci de sélectionner le fichier excel contenant les demandes des périodes initiales, si vous désirez voir les graphes, veuillez re-selectionner le même fichier"), columnspan=3, pady = 10)
  reset.but=tkbutton(tt, text="Remettre les valeurs par défaut", command=reset)
  tkgrid(tklabel(tt,text="Veuillez choisir la période du début"), Debut.entry, pady= 10, padx= 10)
  tkgrid(tklabel(tt,text="Veuillez choisir la valeur d'Alpha"), Alpha.entry, pady= 10, padx= 10)
  tkgrid(reset.but, LissageSimple.but, VisuLissSimplee.but,  pady=10, padx=10)
  tkgrid(Mesure_Precision.but, pady=10, padx=10)
  tkgrid()
  
}
ADDITIONDATAFRAMES_DOUBLE=function(data) {
  x=data.frame(data,Lt=1:length(data[1]),Tt=1:length(data[1]),Pt=1:length(data[1]),Ecart=1:length(data[1]),Ecart_Absolu=1:length(data[1]))
}

Calcul_Lt_Tt_Pt_Init_DOUBLE=function(dataset,x) {
  dataset[x,3]=mean(unlist(dataset[2],use.names = FALSE)[1:x])
  dataset[x,4]=(dataset[x,2]-dataset[1,2])/(dataset[x,1]-dataset[1,1])
  dataset[x,5]=dataset[x,3]+dataset[x,4]
  return(dataset)
}

Calcul_Lt_Tt_Pt_DOUBLE=function(dataset,x,alpha,beta) {
  for (i in 1:(length(unlist(dataset[1],use.names = FALSE))-x)) {
    dataset[i+x,3]=alpha*dataset[i+x-1,2]+(1-alpha)*dataset[i+x-1,5]
    dataset[i+x,4]=beta*(dataset[i+x,3]-dataset[i+x-1,3])+(1-beta)*dataset[i+x-1,4]
    dataset[i+x,5]=dataset[i+x,4]+dataset[i+x,3]
  }
  return(dataset)
}

Calcul_Ecart_EcartAbsolu_DOUBLE=function(dataset,x) {
  for (i in 1:(length(unlist(dataset[1],use.names = FALSE))-x)) {
    dataset[i+x,6]=dataset[i+x,2]-dataset[i+x,5]
    dataset[i+x,7]=abs(dataset[i+x,6])
  }
  return(dataset)
}

Lissage_Double=function(Debut,alpha,beta) {
  library("readxl")
  library("writexl")
  x=read_excel(choose.files())
  x=ADDITIONDATAFRAMES_DOUBLE(x)
  LISSAGEDOUBLE=Calcul_Ecart_EcartAbsolu_DOUBLE(Calcul_Lt_Tt_Pt_DOUBLE(Calcul_Lt_Tt_Pt_Init_DOUBLE(x,Debut),Debut,alpha,beta),Debut)
  write_xlsx(y,"C:\\Users\\msi\\Desktop\\Résultat_Lissage_Double.xlsx")
  return(LISSAGEDOUBLE)
}

Visualisation_Graphique_Lissage_Double=function(x) {
  A=unlist(x[2],use.names = FALSE)
  B=unlist(x[5],use.names = FALSE)
  png(file = "Graphe du lissage double.jpg")
  plot(A,type = "o",col = "red", xlab = "Période", ylab = "Quantité", main = "Prévision sur la demande, Méthode Lissage double")
  lines(B, type = "o", col = "blue")
  dev.interactive()
  dev.off()
}

require(tcltk)
Fenetre_Double=function() {
  
  DebutVariable=tclVar("")
  AlphaVariable=tclVar("")
  BetaVariable=tclVar("")
  tt=tktoplevel()
  tkwm.title(tt,"Prévision de la demande - Lissage double")
  
  Debut.entry=tkentry(tt,textvariable=DebutVariable)
  Alpha.entry=tkentry(tt,textvariable=AlphaVariable)
  Beta.entry=tkentry(tt,textvariable=BetaVariable)
  
  Lissage_Doublee=function() {
    p=Lissage_Double(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable)))
    tkmessageBox(message=paste("Merci de trouver le resultat dans le nouveau excel crée dans le bureau sous le nom de Résultat_Lissage_Double.xlsx, veuillez aussi créer une copie car le document va se faire écraser dans le cas d'un nouveau usage de l'application"))
    return(p)
  }
  LissageDouble.but=tkbutton(tt, text="Start", command=Lissage_Doublee)
  
  
  Visu_Graphique_Double=function() {
    Visualisation_Graphique_Lissage_Double(Lissage_Double(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable))))
    tkmessageBox(message=paste("Merci de trouver le graphe dans le dossier 'Documents', veuillez également créer une copie car l'image va se faire écraser dans le cas d'un nouveau usage de l'application"))
  }
  VisuLissDouble.but=tkbutton(tt, text="Visualiser Les graphes Demandes/Prévisions", command=Visu_Graphique_Double)
  Calcul_CFE_DOUBLEE=function() {
    a=Calcul_MAD_DOUBLE(Lissage_Double(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable))),as.double(tclvalue(DebutVariable)))
    b=Calcul_CFE_DOUBLE(Lissage_Double(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable))),as.double(tclvalue(DebutVariable)))
    c=Calcul_MSE_DOUBLE(Lissage_Double(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable))),as.double(tclvalue(DebutVariable)))
    tkmessageBox(message=paste("Mean Absolute Deviation =",a))
    tkmessageBox(message=paste("Cumulative Forecast Error =",b))
    tkmessageBox(message=paste("Mean Square Error =",c))
  }
  
  Mesure_Precision_double.but=tkbutton(tt, text="Afficher les indicateurs mesurants la précision de la prévision", command=Calcul_CFE_DOUBLEE)
  
  reset=function(){
    tclvalue(DebutVariable)=4
    tclvalue(AlphaVariable)=0.3
    tclvalue(BetaVariable)=0.2
  }
  tkgrid(tklabel(tt,text="Merci de sélectionner le fichier excel contenant les demandes des périodes initiales, si vous désirez voir les graphes, veuillez re-selectionner le même fichier"), columnspan=3, pady = 10)
  reset.but=tkbutton(tt, text="Remettre les valeurs par défaut", command=reset)
  tkgrid(tklabel(tt,text="Veuillez choisir la période du début"), Debut.entry, pady= 10, padx= 10)
  tkgrid(tklabel(tt,text="Veuillez choisir la valeur d'Alpha"), Alpha.entry, pady= 10, padx= 10)
  tkgrid(tklabel(tt,text="Veuillez choisir la valeur de Beta"), Beta.entry, pady= 10, padx= 10)
  tkgrid(reset.but, LissageDouble.but,VisuLissDouble.but, pady=10, padx=10)
  tkgrid(Mesure_Precision_double.but, pady=10, padx=10)
  tkgrid()
  
}

ADDITIONDATAFRAMES=function(data) {
  x=data.frame(data,St=1:length(data[1]),Lt=1:length(data[1]),Tt=1:length(data[1]),Pt=1:length(data[1]),Pt_corrigée=1:length(data[1]),Ecart=1:length(data[1]),Ecart_Absolu=1:length(data[1]))
  
}
Calcul_St_Lt_Tt_Pt_Init=function(dataset,x) {
  for (i in 1:x)
    dataset[i,3]=dataset[i,2]/mean(unlist(dataset[2],use.names = FALSE)[1:x])
  dataset[x,4]=mean(unlist(dataset[2],use.names = FALSE)[1:x])
  dataset[x,5]=(dataset[x,2]-dataset[1,2])/(dataset[x,1]-dataset[1,1])
  dataset[x,6]=dataset[x,5]+dataset[x,4]
  return(dataset)
}

Calcul_St_Lt_Tt_Pt=function(dataset,x,alpha,beta,gamma) {
  for (i in 1:(length(unlist(dataset[1],use.names = FALSE))-x)) {
    dataset[i+x,3]=gamma*(dataset[i+x-1,2]/dataset[i+x-1,4])+(1-gamma)*dataset[i,3]
    dataset[i+x,4]=alpha*(dataset[i+x-1,2]/dataset[i,3])+(1-alpha)*dataset[i+x-1,6]
    a=dataset[i+x,4]
    dataset[i+x,5]=beta*(a-dataset[i+x-1,4])+(1-beta)*dataset[i+x-1,5]
    dataset[i+x,6]=dataset[i+x,5]+dataset[i+x,4]
  }
  return(dataset)
}

Calcul_Ptcorrigé_Ecart_EcartAbsolu=function(dataset,x) {
  for (i in 1:(length(unlist(dataset[1],use.names = FALSE))-x)) {
    dataset[i+x,7]=dataset[i+x,6]*dataset[i,3]
    dataset[i+x,8]=dataset[i+x,2]-dataset[i+x,7]
    dataset[i+x,9]=abs(dataset[i+x,8])
  }
  return(dataset)
}

#library("writexl")
#write_xlsx(x,"C:\\Users\\msi\\Desktop\\test2.xlsx")


Lissage_Triple=function(Debut,alpha,beta,gamme) {
  library("readxl")
  library("writexl")
  x=read_excel(choose.files())
  x=ADDITIONDATAFRAMES(x)
  LISSAGETRIPLE=Calcul_Ptcorrigé_Ecart_EcartAbsolu(Calcul_St_Lt_Tt_Pt(Calcul_St_Lt_Tt_Init(x,Debut),Debut,alpha,beta,gamme),Debut)
  write_xlsx(y,"C:\\Users\\msi\\Desktop\\Méthode du Lissage Triple.xlsx")
  return(LISSAGETRIPLE)
}

Visualisation_Graphique_Lissage_Triple=function(x) {
  A=unlist(x[2],use.names = FALSE)
  B=unlist(x[7],use.names = FALSE)
  png(file = "Graphe du lissage triple.jpg")
  plot(A,type = "o",col = "red", xlab = "Période", ylab = "Quantité", main = "Prévision sur la demande, Méthode Lissage triple")
  lines(B, type = "o", col = "blue")
  dev.interactive()
  dev.off()
}

require(tcltk)
Fenetre=function() {
  
  DebutVariable=tclVar("")
  AlphaVariable=tclVar("")
  BetaVariable=tclVar("")
  GammaVariable=tclVar("")
  
  tt=tktoplevel()
  tkwm.title(tt,"Prévision de la demande - Lissage triple")
  
  Debut.entry=tkentry(tt,textvariable=DebutVariable)
  Alpha.entry=tkentry(tt,textvariable=AlphaVariable)
  Beta.entry=tkentry(tt,textvariable=BetaVariable)
  Gamma.entry=tkentry(tt,textvariable=GammaVariable)
  
  Lissage_Triplee=function() {
    p=Lissage_Triple(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable)),as.double(tclvalue(GammaVariable)))
    tkmessageBox(message=paste("Merci de trouver le resultat dans le nouveau excel crée dans le bureau sous le nom de test6.xlsx, veuillez aussi créer une copie car le document va se faire écraser dans le cas d'un nouveau usage de l'application"))
    return(p)
  }
  Lissage.but=tkbutton(tt, text="Start", command=Lissage_Triplee)
  Visu_Graphique_Triple=function() {
    Visualisation_Graphique_Lissage_Triple(Lissage_Triple(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable)),as.double(tclvalue(GammaVariable))))
    tkmessageBox(message=paste("Merci de trouver le graphe dans le dossier 'Documents', veuillez également créer une copie car l'image va se faire écraser dans le cas d'un nouveau usage de l'application"))
  }
  VisuLissTripl.but=tkbutton(tt, text="Visualiser Les graphes Demandes/Prévisions", command=Visu_Graphique_Triple)
  Calcul_CFE_TRIPLEE=function() {
    d=Calcul_MAD_TRIPLE(Lissage_Triple(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable)),as.double(tclvalue(GammaVariable))),as.double(tclvalue(DebutVariable)))
    e=Calcul_CFE_TRIPLE(Lissage_Triple(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable)),as.double(tclvalue(GammaVariable))),as.double(tclvalue(DebutVariable)))
    f=Calcul_MSE_TRIPLE(Lissage_Triple(as.double(tclvalue(DebutVariable)),as.double(tclvalue(AlphaVariable)),as.double(tclvalue(BetaVariable)),as.double(tclvalue(GammaVariable))),as.double(tclvalue(DebutVariable)))
    tkmessageBox(message=paste("Mean Absolute Deviation =",d))
    tkmessageBox(message=paste("Cumulative Forecast Error =",e))
    tkmessageBox(message=paste("Mean Square Error =",f))
  }
  
  Mesure_Precision_triple.but=tkbutton(tt, text="Afficher les indicateurs mesurants la précision de la prévision", command=Calcul_CFE_TRIPLEE)
  
  reset=function(){
    tclvalue(DebutVariable)=4
    tclvalue(AlphaVariable)=0.3
    tclvalue(BetaVariable)=0.2
    tclvalue(GammaVariable)=0.1
  }
  reset.but=tkbutton(tt, text="Remettre les valeurs par défaut", command=reset)
  
  tkgrid(tklabel(tt,text="Merci de sélectionner le fichier excel contenant les demandes des périodes initiales, si vous désirez voir les graphes, veuillez re-selectionner le même fichier"), columnspan=3, pady = 10)
  tkgrid(tklabel(tt,text="Veuillez choisir la période du début"), Debut.entry, pady= 10, padx= 10)
  tkgrid(tklabel(tt,text="Veuillez choisir la valeur d'Alpha"), Alpha.entry, pady= 10, padx= 10)
  tkgrid(tklabel(tt,text="Veuillez choisir la valeur de Beta"), Beta.entry, pady= 10, padx= 10)
  tkgrid(tklabel(tt,text="Veuillez choisir la valeur de Gamma"), Gamma.entry, pady= 10, padx= 10)
  tkgrid(reset.but, Lissage.but,VisuLissTripl.but, pady=10, padx=10)
  tkgrid(Mesure_Precision_triple.but, pady=10, padx=10)
  tkgrid()
  
}

require(tcltk)
Fenetre_Principale=function() {
  tt <- tktoplevel()
  tkwm.title(tt,"Application GOP - Prévision de la demande")
  Triplee<- function() {
    Fenetre()
    tkmessageBox(message=paste("fichier chargé avec succès"))
  }
  Triplee.but <- tkbutton(tt, text="Méthode du lissage triple", command=Triplee)
  Doublee<- function() {
    Fenetre_Double()
    tkmessageBox(message=paste("fichier chargé avec succès"))
  }
  Doublee.but <- tkbutton(tt, text="Méthode du lissage double", command=Doublee)
  Simplee<- function() {
    Fenetre_Simple()
    tkmessageBox(message=paste("fichier chargé avec succès"))
  }
  simplee.but <- tkbutton(tt, text="Méthode du lissage simple", command=Simplee)
  
  tkgrid(tklabel(tt,text="Bienvenue"), columnspan=3, pady = 10)
  tkgrid(tklabel(tt,text="Merci de choisir la méthode que vous désirez appliquer"), columnspan=3, pady = 10)
  tkgrid(simplee.but, Doublee.but, Triplee.but, pady= 10, padx= 10)
}




