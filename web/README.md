# Version Web — RCS Madagascar

Application web Streamlit du scraper RCS Madagascar. Marche sur n'importe quel navigateur (Mac, Windows, Linux, tablette, mobile).

L'utilisateur final n'a qu'une URL à ouvrir, aucune installation.

## Tester en local

Depuis ce dossier `web/` :

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

L'app s'ouvre dans le navigateur à http://localhost:8501.

## Déployer sur Streamlit Community Cloud (gratuit)

### Étape 1 — Mettre le projet sur GitHub

1. Créer un compte GitHub sur https://github.com (gratuit)
2. Créer un nouveau repo public (par exemple `rcs-madagascar`)
3. Pousser le dossier `rcs/` complet (avec `rcs_scraper.py` à la racine et le dossier `web/` à l'intérieur)

### Étape 2 — Connecter Streamlit Cloud

1. Aller sur https://share.streamlit.io et se connecter avec GitHub
2. Cliquer **« New app »**
3. Configurer :
   - **Repository** : ton-utilisateur/rcs-madagascar
   - **Branch** : main
   - **Main file path** : `web/streamlit_app.py`
4. Cliquer **« Deploy »**

L'app sera live en 2-3 minutes à une URL du type :
```
https://rcs-madagascar.streamlit.app
```

Tu peux partager cette URL à qui tu veux. À chaque push sur GitHub, l'app se met à jour automatiquement.

## Structure du dossier

```
web/
├── streamlit_app.py        Application Streamlit (interface web)
├── requirements.txt        Dépendances Python (streamlit + scraper)
├── .streamlit/
│   └── config.toml         Thème et configuration serveur
└── README.md               Ce fichier
```

L'app importe la logique de scraping depuis `../rcs_scraper.py` (aucune duplication de code).

## Limites du tier gratuit Streamlit Cloud

- 1 GB de RAM (suffisant pour la plupart des extractions)
- L'app se met en veille après ~15 min d'inactivité (premier visiteur attend ~10-30s pour le réveil)
- Code visible sur GitHub (repo public requis pour le gratuit)
- 1 app gratuite par utilisateur (limite récente)

Pour passer au privé ou avoir plus de ressources : abonnement Streamlit Cloud payant, ou alternatives gratuites (Hugging Face Spaces, Render, Fly.io).
