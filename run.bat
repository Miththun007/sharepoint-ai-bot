@echo off
title SharePoint AI Bot Web Service
cd /d "%~dp0"
echo ==========================================================
echo Starting SharePoint AI Assistant Web Server...
echo Site: https://veeleaddev.sharepoint.com/sites/AISearchDev
echo Model: Azure OpenAI (gpt-4o-mini)
echo ==========================================================
python app.py
pause
