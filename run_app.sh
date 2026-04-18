#!/bin/bash
echo "╔══════════════════════════════════════════════════╗"
echo "║   CFG Bank – Système de Valorisation OPCVM       ║"
echo "╚══════════════════════════════════════════════════╝"
echo ""
echo "Installation des dépendances..."
pip install -r requirements.txt -q
echo ""
echo "Démarrage de l'application..."
echo "L'application sera accessible sur: http://localhost:8501"
echo ""
streamlit run app.py --server.port 8501 --browser.gatherUsageStats false
