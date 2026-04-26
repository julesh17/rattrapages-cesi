# 🎓 Application Streamlit – Gestion des rattrapages étudiants CESI

## 🚀 Accès à l’application

L’application est disponible en ligne (aucune installation nécessaire) :
👉 https://rattrapages-cesi.streamlit.app

---

## 📌 Présentation

Cette application permet de :

* Analyser les notes des étudiants
* Identifier automatiquement les rattrapages (C, D, ABS)
* Gérer les compensations par UE (Unités d’Enseignement)
* Générer des mails de convocation personnalisés
* Construire des créneaux d’examen sans conflit
* Exporter les résultats en Excel

---

## 🧩 Fonctionnalités principales

### 1. Import des données

* Fichier de notes (.xlsx) obligatoire
* Fichier référentiel RN (.xlsx) optionnel (pour les compensations)

### 2. Configuration

* Choix du semestre
* Exclusion de matières (ex : Global exam)
* Gestion des absences (cellules vides → ABS)

### 3. Analyse des résultats

* Tableau interactif des notes
* Filtres par mention (A, B, C, D)
* Filtre par groupe (admis / rattrapage)

### 4. Compensations UE

* Calcul automatique des moyennes pondérées
* Détection des UE validées ou compensées
* Visualisation détaillée par étudiant

### 5. Récapitulatif par matière

* Nombre d’étudiants en rattrapage
* Distinction C / D / ABS
* Identification des compensations

### 6. Génération de mails

* Mails individuels personnalisés
* Mail récapitulatif pour la classe
* Option tutoiement / vouvoiement
* Copie rapide ou téléchargement

### 7. Créneaux parallèles

* Proposition de groupes de matières compatibles
* Évite les conflits d’étudiants
* Matrice de compatibilité complète

### 8. Export Excel

Contient plusieurs onglets :

* **Vue complète** : tous les étudiants
* **Mentions (filtrés)** : selon les filtres
* **Rattrapages** : synthèse globale
* **Résultats UE** : détails des compensations

---

## ⚙️ Utilisation

1. Ouvrir l’application en ligne
2. Importer le fichier de notes issu de Scholaris
3. (Optionnel) Importer le référentiel RN issu d'une requête Bora
4. Configurer les options
5. Explorer les résultats dans les onglets
6. Générer les mails ou exporter les données

---

## 💻 Lancement en local (optionnel)

Si besoin, vous pouvez exécuter l’application localement :

```bash
pip install streamlit pandas openpyxl
streamlit run rattrapages_app_v4.py
```

---

## 📁 Dépendances

* streamlit
* pandas
* openpyxl

---

## 📝 Remarques

* Les mentions sont converties en valeurs pour les calculs :
  A=5, B=4, C=2, D=1
* Les ABS sont traités comme D dans les moyennes UE
* Les compensations s’appliquent uniquement si l’UE est validée

---

## 👥 Public cible

* Enseignants Responsables Pédagogiques
* Gestionnaires académiques

---

## ✨ Objectif

Simplifier et automatiser la gestion des rattrapages étudiants, tout en réduisant les erreurs et le temps de traitement.
