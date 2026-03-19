# 🎓 Gestion des rattrapages

Application Streamlit pour visualiser les résultats d'étudiants, filtrer par mention et générer des mails de convocation aux rattrapages.

## Fonctionnalités

- **Import Excel** : accepte tout fichier `.xlsx` avec une colonne `Personne` et des colonnes `Eval - ...`
- **Tableau colorisé** : badges A (vert) / B (bleu) / C (jaune) / D (rouge)
- **Filtres** : par mention individuelle et par groupe (A/B ou C/D)
- **Export Excel** : tableau filtré avec couleurs
- **Mails de convocation** : générés automatiquement, éditables, avec choix tutoiement/vouvoiement

## Lancer en local

```bash
pip install -r requirements.txt
streamlit run rattrapages_app.py
```

## Déployer sur Streamlit Cloud

1. Pusher ce repo sur GitHub
2. Aller sur [share.streamlit.io](https://share.streamlit.io)
3. Connecter le repo et sélectionner `rattrapages_app.py`
4. Cliquer sur **Deploy**

## Format du fichier Excel attendu

| Personne | Eval - Matière 1 | Eval - Matière 2 | ... |
|---|---|---|---|
| NOM, Prénom | A | C | ... |

- La colonne `Personne` doit être au format `NOM, Prénom`
- Les colonnes d'évaluation doivent commencer par `Eval`
- Les valeurs acceptées sont : `A`, `B`, `C`, `D`
