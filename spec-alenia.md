# ALENIA — Dossier de Spécifications Fonctionnelles & Techniques
**Version :** 1.2 (Consolidée)  
**Date :** Octobre 2023  
**Statut :** Validé pour Développement MVP  
**Auteur :** Lead Business Analyst / Product Analyst  

---

## TABLE DES MATIÈRES
1. [Executive Summary](#1-executive-summary)
2. [Vision Produit](#2-vision-produit)
3. [Business Goals & Métriques de Succès](#3-business-goals--métriques-de-succès)
4. [Principes Directeurs (Product Principles)](#4-principes-directeurs-product-principles)
5. [Personas](#5-personas)
6. [Stakeholders](#6-stakeholders)
7. [Scope (Périmètre MVP)](#7-scope-périmètre-mvp)
8. [Out of Scope](#8-out-of-scope)
9. [Hypothèses](#9-hypothèses)
10. [Contraintes](#10-contraintes)
11. [Architecture Fonctionnelle & Technique Locale](#11-architecture-fonctionnelle--technique-locale)
12. [Parcours Utilisateurs (User Journeys)](#12-parcours-utilisateurs-user-journeys)
13. [Use Cases Détaillés](#13-use-cases-détaillés)
14. [Exigences Fonctionnelles (FR)](#14-exigences-fonctionnelles-fr)
15. [Exigences Non-Fonctionnelles (NFR)](#15-exigences-non-fonctionnelles-nfr)
16. [Exigences IA & LLM (AI)](#16-exigences-ia--llm-ai)
17. [Exigences de Sécurité (SEC)](#17-exigences-de-sécurité-sec)
18. [Exigences de Confidentialité & RGPD (PRIV)](#18-exigences-de-confidentialité--rgpd-priv)
19. [Moteur DLP & Tokenisation (DLP)](#19-moteur-dlp--tokenisation-dlp)
20. [Modèle de Rôles & RBAC](#20-modèle-de-rôles--rbac)
21. [Module d'Administration (ADM)](#21-module-dadministration-adm)
22. [Module Analytics & Confidentialité](#22-module-analytics--confidentialité)
23. [Modèle de Données (Data Model Local)](#23-modèle-de-données-data-model-local)
24. [Architecture des Agents](#24-architecture-des-agents)
25. [LLM Gateway & Mécanisme de Fallback](#25-llm-gateway--mécanisme-de-fallback)
26. [Intégration Microsoft Word (Web Add-in)](#26-intégration-microsoft-word-web-add-in)
27. [Matrice de Compatibilité](#27-matrice-de-compatibilité)
28. [Spécifications UX/UI](#28-spécifications-uxui)
29. [Wireframes Textuels Complets](#29-wireframes-textuels-complets)
30. [Gestion des Erreurs (Error Handling)](#30-gestion-des-erreurs-error-handling)
31. [Modes Hors-Ligne & Dégradé](#31-modes-hors-ligne--dégradé)
32. [Exigences de Performance (PERF)](#32-exigences-de-performance-perf)
33. [Observabilité & Journalisation (OBS)](#33-observabilité--journalisation-obs)
34. [Threat Model (Modèle de Menaces)](#34-threat-model-modèle-de-menaces)
35. [Cas Limites (Edge Cases)](#35-cas-limites-edge-cases)
36. [User Stories (US)](#36-user-stories-us)
37. [Critères d'Acceptation (Gherkin)](#37-critères-dacceptation-gherkin)
38. [Stratégie de Test Globale](#38-stratégie-de-test-globale)
39. [Stratégie UAT & Cahier de Recette Détaillé](#39-stratégie-uat--cahier-de-recette-détaillé)
40. [Définition du MVP](#40-définition-du-mvp)
41. [Roadmap Produit (MVP, V1.5, V2)](#41-roadmap-produit-mvp-v15-v2)
42. [Registre des Risques & Plans de Mitigation](#42-registre-des-risques--plans-de-mitigation)
43. [Arbre des Dépendances](#43-arbre-des-dépendances)
44. [Registre des Questions Ouvertes & Résolutions](#44-registre-des-questions-ouvertes--résolutions)
45. [Decision Log Consolidé (DEC-001 à DEC-025)](#45-decision-log-consolidé-dec-001-à-dec-025)
46. [Matrice de Traçabilité Bout-en-Bout](#46-matrice-de-traçabilité-bout-en-bout)

---

## 1. EXECUTIVE SUMMARY

**ALENIA** est un assistant rédactionnel professionnel basé sur l'intelligence artificielle générative, conçu pour assister les experts, économistes, géopolitologues et analystes du secteur public et privé dans la rédaction de documents à haute valeur ajoutée.

Au stade MVP, ALENIA est déployé sur **Windows 10/11** pour un déploiement pilote de **10 postes**. L'intégration au flux de travail de l'utilisateur repose sur un double canal complémentaire :
1. Une **Taskpane Microsoft Word** (JS Web Add-in / Office.js sideloadé) servant d'ancrage permanent et d'historique.
2. Un **Overlay universel léger** (WPF / .NET 8) déclenché par raccourci global (`Ctrl+Shift+N`) affichant le calcul de différence visuelle (DIFF) interactif.

Afin de garantir une sécurité maximale et une conformité stricte avec les exigences d'entreprise (documents C2 - internes sans diffusion publique), ALENIA fonctionne selon une architecture **100% locale** : un serveur ASP.NET Core Kestrel in-process tourne sur le poste (`127.0.0.1:5284`), assurant la médiation entre l'Add-in Word, l'Overlay, le moteur DLP (anonymisation/tokenisation réversible), la base SQLite locale, et la passerelle d'accès direct et sécurisé vers les API des modèles de langage (Claude latest par défaut, Gemini en repli).

---

## 2. VISION PRODUIT

ALENIA n'est pas un rédacteur autonome, mais un **copilote d'amplification intellectuelle et stylistique**. Il garantit la totale maîtrise de l'utilisateur sur son texte (*Human-in-the-loop* absolu). Il permet aux rédacteurs d'atteindre une rigueur formelle, stylistique, conceptuelle et quantitative sans jamais modifier le document sous-jacent de manière silencieuse ou non sollicitée.

---

## 3. BUSINESS GOALS & MÉTRIQUES DE SUCCÈS

### 3.1 Objectifs Métier (Business Goals)
- **BG-01 :** Réduire de 40 % le temps passé à la relecture orthotypographique, stylistique et de cohérence des notes d'analyse.
- **BG-02 :** Garantir une conformité éditoriale institutionnelle et bilingue (FR ↔ EN UK/US) sans recours systématique à des relecteurs externes.
- **BG-03 :** Sécuriser la fuite d'informations sensibles (DLP) lors de l'interaction avec des LLMs commerciaux.
- **BG-04 :** Assurer une adoption fluide sans friction d'onboarding (< 5 minutes de prise en main).

### 3.2 KPIs de Succès (Pilote 10 postes)
- **Taux d'acceptation des suggestions :** $\ge 70\%$ des propositions acceptées ou acceptées après modification mineure.
- **Fréquence d'utilisation :** $\ge 5$ interactions par utilisateur actif par jour.
- **Taux de faux positifs DLP :** $< 5\%$ sur les entités nommées courantes.
- **Stabilité applicative :** Zéro crash bloquant sur Word ou sur le processus ALENIA sur une durée de 30 jours continus.

---

## 4. PRINCIPES DIRECTEURS (PRODUCT PRINCIPES)

| ID | Nom | Description opérationnelle |
|:---|:---|:---|
| **P-001** | **Security by Design** | Secrets stockés via DPAPI Windows, binding réseau local strict sur `127.0.0.1`, validation des origines. |
| **P-002** | **Privacy by Design** | Aucune conservation persistante du texte des documents sur disque ou serveur distant. Traitement éphémère en RAM. |
| **P-003** | **Generative AI by Design** | Délimitation stricte entre instructions système et données documentaires (anti-prompt injection). |
| **P-004** | **Human-in-the-loop** | Aucune modification silencieuse. Chaque modification requiert un clic "Accepter". |
| **P-005** | **Model Agnostic** | Passerelle modulaire permettant de basculer d'Anthropic à Google ou à un LLM local sans impacter le cœur. |
| **P-006** | **Modular Architecture** | Agents découplés configurés par descripteurs JSON standardisés. |
| **P-007** | **Extensibility by Design** | Nouveaux agents paramétrables par l'administrateur sans recompilation. |
| **P-008** | **Low Resource Consumption**| Consommation RAM du processus ALENIA en tâche de fond $< 150\text{ Mo}$. |
| **P-009** | **Non-Intrusive UX** | L'Overlay n'apparaît que sur demande expresse et libère le focus instantanément. |
| **P-010** | **Interoperability** | Standardisation des échanges via API HTTP/WebSocket et formats HTML/JSON. |
| **P-011** | **Minimal Collection** | Analytics strictement locales au MVP, anonymisées et sans données textuelles. |
| **P-012** | **Auditability** | Traçabilité des actions, volumes de tokens et temps de réponse sans capture du contenu. |
| **P-013** | **Progressive Evolution** | Architecture prête pour le portage macOS et le backend centralisé en V2. |
| **P-014** | **Preservation of Workflow**| Respect de la sélection active, de l'historique d'annulation Word (Undo/Redo). |
| **P-015** | **Graceful Degradation** | Messages clairs et bascule vers le presse-papiers si un composant échoue. |

---

## 5. PERSONAS

### Persona 1 : Paul — Économiste Senior (48 ans)
- **Rôle :** Rédige des notes de conjoncture macroéconomique, des synthèses trimestrielles de 2 à 5 pages sous Word.
- **Besoins :** Correction orthotypographique exigeante (espaces insécables, ponctuation financière), vérification de cohérence des chiffres et pourcentages, traduction institutionnelle vers l'anglais UK.
- **Pain points :** Perte de temps sur les relectures de forme, peur d'inverser des ratios ou des dates, frustration face aux outils grand public qui altèrent sa mise en page.

### Persona 2 : Sarah — Analyste Géopolitique (34 ans)
- **Rôle :** Produit des analyses prospectives, des notes d'alerte et des mémos stratégiques.
- **Besoins :** Reformulation pour adapter le niveau de langage (décideur politique vs expert technique), respect strict des acronymes internationaux, définition contextuelle de concepts émergents.
- **Pain points :** Risque de fuite de noms de projets ou de personnalités diplomatiques lors de l'utilisation d'outils cloud non autorisés.

### Persona 3 : Julie — Administratrice IT / Responsable Sécurité Déploiement (41 ans)
- **Rôle :** Gestionnaire du parc applicatif et garante de la conformité poste de travail.
- **Besoins :** Déploiement automatisable (MSI), contrôle des flux réseau, configuration simple des clés API et règles DLP, zéro impact sur les performances de Word.
- **Pain points :** Peur des plug-ins lourds qui font crasher Office, gestion complexe des serveurs d'infrastructure.

---

## 6. STAKEHOLDERS

- **Direction Générale / Métier :** Exige une élévation globale de la qualité rédactionnelle et une réduction des délais de livraison des notes.
- **Utilisateurs Analystes :** Exigent un outil ergonomique, ultra-rapide et respectueux de leur style.
- **DSI & RSSI :** Exigent l'étanchéité des données (DLP), le contrôle des flux sortants (Proxy corporate), et l'absence d'infrastructure serveur lourde pour le pilote.
- **Équipe de Développement :** Exige des spécifications exhaustives, sans ambiguïté technique, facilitant la réalisation de tests automatisés et UAT sur VM distante.

---

## 7. SCOPE (PÉRIMÈTRE MVP)

- **OS :** Windows 10 (version 21H2+) et Windows 11 (x64).
- **Application Hôte :** Microsoft Word (Microsoft 365 Desktop & Office LTSC 2021).
- **Interfaces :**
  - Web Add-in Word avec Taskpane latérale (HTML/JS/Office.js).
  - Overlay flottant natif (WPF .NET 8) avec affichage DIFF interactif.
  - Console d'administration Web locale (`http://127.0.0.1:5284/admin`).
  - Menu System Tray (barre des tâches).
- **Agents Natifs (5) :**
  1. Correction (orthographe, grammaire, typographie, style).
  2. Reformulation (modes prédéfinis + consigne libre).
  3. Traduction bilingue (FR ↔ EN US / EN UK).
  4. Cohérence Data (moteur numérique déterministe + LLM).
  5. Définition / Explication (analyse de concept, affichage overlay).
- **Moteur DLP :** Regex + NER basique + anonymisation/tokenisation réversible locale.
- **Connectivité LLM :** Passerelle locale multi-fournisseur avec Anthropic (Claude 3.5 Sonnet / latest) en primaire et Google Gemini en repli automatique. Option BYOK par poste.
- **Gestion des Licences :** Fichier de licence local chiffré pour 10 postes liés aux identifiants matériels (Hardware ID).
- **Déploiement :** Installateur MSI client + Sideload du fichier `manifest.xml` pour l'Add-in Word.

---

## 8. OUT OF SCOPE

- **Mode "ALL" (traitement du document complet sans sélection) :** Reporté en V1.5.
- **Support des autres applications hôtes (PowerPoint, Outlook, Excel, Notepad) :** Reporté en V1.5.
- **LLM 100% local sur carte graphique (Ollama / Llama.cpp) :** Reporté en V1.5.
- **Serveur backend centralisé Cloud ALENIA :** Reporté en V2.
- **Portage macOS (macOS Desktop + Word Mac) :** Reporté en V2.
- **Module OCR / Vision (analyse d'images et graphiques) :** Reporté en V2.
- **Détection automatisée des biais idéologiques et analyse de neutralité :** Reporté en V2.
- **Intégration SSO d'entreprise (Azure AD / Entra ID / Okta) :** Reporté en V1.5.
- **Plateformes mobiles (Android / iOS) :** Définitivement exclu.

---

## 9. HYPOTHÈSES

- **H-01 :** Les postes utilisateurs disposent du runtime WebView2 (installé par défaut sur Windows 10/11 et Office 365/2021) permettant d'exécuter les Web Add-ins Office.js.
- **H-02 :** La politique de sécurité locale du poste autorise l'ouverture d'un port TCP d'écoute sur l'interface de boucle locale `127.0.0.1`.
- **H-03 :** Les flux HTTPS sortants vers `api.anthropic.com` et `generativelanguage.googleapis.com` sont autorisés en direct ou à travers le serveur proxy de l'entreprise.
- **H-04 :** L'utilisateur travaille par sélection de blocs textuels homogènes (paragraphe, groupe de paragraphes, tableau textuel de 1 à 1 000 mots).
- **H-05 :** Les administrateurs locaux disposent des droits nécessaires pour exécuter le programme d'installation MSI sur les 10 machines pilotes.

---

## 10. CONTRAINTES

- **C-01 (Architecture) :** Aucun serveur cloud ALENIA intermédiaire au MVP. L'exécutable desktop héberge son propre serveur Kestrel local.
- **C-02 (Réseau d'Entreprise) :** Gestion impérative des proxies d'entreprise (authentifiés NTLM/Basic ou anonymes) configurables localement.
- **C-03 (Format Document) :** Préservation obligatoire des formats de base Word (gras, italique, souligné, couleur, polices, hyperliens) lors de la réinjection via Office.js `insertHtml`.
- **C-04 (Sécurité Noyau) :** Respect strict de l'UIPI (User Interface Privilege Isolation) de Windows : l'overlay ne doit pas chercher à injecter des inputs dans des processus d'intégrité supérieure.
- **C-05 (Environnement de Test) :** L'environnement de développement ne disposant pas de Windows, l'ensemble des livrables doit intégrer un protocole d'automatisation et de documentation permettant une exécution isolée des tests sur VM Windows distante.

---

## 11. ARCHITECTURE FONCTIONNELLE & TECHNIQUE LOCALE

```
┌──────────────────────────────────────────────────────────────────────────────┐
│ POSTE DE TRAVAIL WINDOWS 10/11 (Architecture Locale MVP)                    │
│                                                                              │
│  ┌───────────────────────────────┐     ┌──────────────────────────────────┐  │
│  │   MICROSOFT WORD             │     │   ALENIA WPF OVERLAY             │  │
│  │                               │     │   (Fenêtre native TopMost)       │  │
│  │  ┌─────────────────────────┐  │     │                                  │  │
│  │  │ JS Web Add-in           │  │     │  - Affichage Texte Source        │  │
│  │  │ (Office.js Taskpane)    │  │     │  - Sélecteur d'Agents            │  │
│  │  │                         │  │     │  - Visualiseur DIFF Interactif   │  │
│  │  │ - Boutons d'action      │  │     │  - Actions : Accepter / Rejeter  │  │
│  │  │ - État connexion WS     │  │     │  - Édition manuelle / Consigne   │  │
│  │  │ - Historique récent     │  │     └────────────────┬─────────────────┘  │
│  │  └────────────┬────────────┘  │                      │ Interop IPC        │
│  └───────────────┼───────────────┘                      │ In-Process         │
│                  │ WebSocket / HTTP                     │                    │
│                  │ (ws://127.0.0.1:5284/ws/addin)       ▼                    │
│                  │ (http://127.0.0.1:5284/api/*)  ┌───────────────────────┐  │
│                  ▼                                │ ALENIA CORE (.NET 8)  │  │
│  ┌────────────────────────────────────────────────┤                       │  │
│  │ KESTREL WEB SERVER (In-Process Desktop Client) │ - Win32 Keyboard Hook │  │
│  │                                                │   (Ctrl+Shift+N)      │  │
│  │  ┌─────────────────┐    ┌───────────────────┐  │ - UI Automation Capture│ │
│  │  │ API Endpoints   │    │ SignalR Hub       │  │ - Licence Manager     │  │
│  │  │ (/capture,      │    │ (/ws/addin)       │  │ - System Tray Manager │  │
│  │  │  /reinject)     │    │                   │  └───────────┬───────────┘  │
│  │  └────────┬────────┘    └─────────┬─────────┘              │              │
│  └───────────┼───────────────────────┼────────────────────────┘              │
│              ▼                       ▼                                       │
│  ┌────────────────────────────────────────────────────────────────────────┐  │
│  │ PIPELINE DE TRAITEMENT DES AGENTS                                      │  │
│  │                                                                        │  │
│  │  ┌─────────────────┐    ┌─────────────────┐    ┌────────────────────┐  │  │
│  │  │ 1. DLP Engine   │───►│ 2. Agent Engine │───►│ 3. LLM Gateway     │  │  │
│  │  │ (Regex + NER +  │    │ (Prompts,       │    │ (Anthropic Client /│  │  │
│  │  │  Tokenisation)  │    │  Profils,       │    │  Gemini Fallback / │  │  │
│  │  │                 │    │  Règles)        │    │  Proxy Handler)    │  │  │
│  │  └────────┬────────┘    └─────────────────┘    └──────────┬─────────┘  │  │
│  │           ▲                                               │            │  │
│  │           │ Restauration inverse (Détokenisation)         │            │  │
│  │           └───────────────────────────────────────────────┘            │  │
│  └───────────────────────────────────┬────────────────────────────────────┘  │
│                                      │                                       │
│  ┌───────────────────────────────────┴────────────────────────────────────┐  │
│  │ COUCHE DE PERSISTANCE & CONFIGURATION LOCALE                           │  │
│  │                                                                        │  │
│  │  - SQLite DB locale : Profils, Glossaires, Règles DLP, Analytics, Logs  │  │
│  │  - Windows DPAPI : Clés d'API chiffrées, Clé de Licence Machine        │  │
│  └───────────────────────────────────┬────────────────────────────────────┘  │
│                                      │                                       │
│  ┌───────────────────────────────────┴────────────────────────────────────┐  │
│  │ CONSOLE D'ADMINISTRATION LOCALE EMBARQUÉE                              │  │
│  │ Web UI servie sur http://127.0.0.1:5284/admin                          │  │
│  └────────────────────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────┼───────────────────────────────────────┘
                                       │ HTTPS Sortant Sécurisé (TLS 1.3)
                                       │ (Direct ou via Proxy Corporate)
                                       ▼
                     ┌───────────────────────────────────┐
                     │ FOURNISSEURS LLM EXTERNES (CLOUD) │
                     │                                   │
                     │  - Anthropic API (Claude 3.5)     │
                     │  - Google Gemini Pro API (Repli)  │
                     │  - Endpoint BYOK personnalisé     │
                     └───────────────────────────────────┘
```

---

## 12. PARCOURS UTILISATEURS (USER JOURNEYS)

### 12.1 Parcours 1 : Relecture & Correction Stylistique via Raccourci Global
1. **Contexte :** L'économiste termine la rédaction d'une note sur l'inflation dans Word.
2. **Déclenchement :** Il sélectionne trois paragraphes contenant des données et du texte enrichi, puis appuie sur `Ctrl+Shift+N`.
3. **Capture :** Le hook global ALENIA capture l'événement. Le client desktop ordonne au Web Add-in via WebSocket d'extraire la sélection HTML (`Office.js getSelectedDataAsync(CoercionType.Html)`).
4. **Affichage :** L'overlay ALENIA s'ouvre à proximité immédiate du curseur, affichant le texte brut extrait tout en conservant l'arbre DOM sous-jacent.
5. **Sélection Agent :** L'utilisateur clique sur **"Corriger"** (Profil actif : *Note Ministérielle*).
6. **Sécurité :** Le moteur DLP local analyse le texte. Deux noms propres de projets internes sont anonymisés (`Projet Helios` $\rightarrow$ `PROJ_001`).
7. **Traitement :** Le LLM Gateway envoie la charge utile sécurisée à Claude 3.5 Sonnet avec le prompt système institutionnel.
8. **Résolution & Diff :** La réponse est reçue, les tokens DLP sont réinjectés pour rétablir les vrais termes. L'Overlay affiche le **DIFF interactif** (rouge barré pour les suppressions, vert souligné pour les ajouts, jaune pour les modifications typographiques).
9. **Validation :** L'utilisateur inspecte les modifications, effectue un ajustement manuel d'un mot dans le visualiseur, puis clique sur **"Accepter"** (ou `Enter`).
10. **Réinjection :** L'Overlay transmet le fragment HTML corrigé au Web Add-in qui exécute `range.insertHtml(html, "Replace")`.
11. **Finalisation :** Le texte dans Word est mis à jour instantanément avec ses polices, couleurs et puces intactes. L'overlay se referme. L'utilisateur peut faire `Ctrl+Z` dans Word s'il souhaite annuler.

### 12.2 Parcours 2 : Traduction Sécurisée avec Entités Sensibles via la Taskpane
1. **Contexte :** L'analyste souhaite traduire un mémo stratégique du français vers l'anglais UK.
2. **Interaction :** L'analyste clique sur l'onglet ALENIA dans le ruban Word pour afficher la Taskpane latérale.
3. **Sélection :** Le texte sélectionné dans Word est automatiquement détecté par la Taskpane.
4. **Action :** L'analyste sélectionne l'agent **"Traduire"**, cible : **Anglais (UK)**, avec le glossaire "Macroéconomie/Finance" activé.
5. **DLP & Tokenisation :** Les montants financiers confidentiels (`45,8 M€`) et identifiants de diplomates sont masqués localement.
6. **Exécution :** La passerelle traite la requête. L'Overlay s'ouvre pour présenter le résultat sous forme de panneau côte-à-côte avec surlignage des termes issus du glossaire métier.
7. **Validation :** L'analyste clique sur "Accepter". Le texte anglais remplace la sélection française dans Word, en appliquant les conventions typographiques anglaises (espaces, virgules décimales).

---

## 13. USE CASES DÉTAILLÉS

### UC-01 : Correction Orthotypographique & Stylistique
- **Acteur Principal :** Utilisateur Word.
- **Préconditions :** Word actif, ALENIA Desktop en cours d'exécution, texte sélectionné.
- **Déclencheur :** Raccourci `Ctrl+Shift+N` ou clic Taskpane.
- **Scénario Nominal :**
  1. Le système capture la sélection active et le formatage HTML.
  2. Le système présente l'overlay avec l'agent "Correction" présélectionné.
  3. L'utilisateur valide l'analyse.
  4. Le système applique les règles du profil éditorial choisi (ex: institutionnel).
  5. Le système présente les modifications sous forme de Diff granulaire (niveau mot/caractère).
  6. L'utilisateur accepte la proposition.
  7. Le document Word est mis à jour avec préservation du style.
- **Scénarios Alternatifs :**
  - *4a. Pas de connectivité Internet :* La passerelle tente le repli vers Gemini. Si échec, affichage d'un message d'erreur : "Service de traitement indisponible. Vérifiez votre connexion proxy."
  - *6a. Rejet par l'utilisateur :* L'overlay se ferme sans altérer le document Word.
  - *6b. Modification manuelle :* L'utilisateur modifie directement le texte dans l'éditeur de l'overlay avant acceptation ; le texte modifié est réinjecté.

### UC-02 : Contrôle Déterministe de Cohérence Numérique
- **Acteur Principal :** Analyste Économiste.
- **Scénario Nominal :**
  1. L'utilisateur sélectionne un paragraphe contenant des calculs, pourcentages et totaux.
  2. L'utilisateur choisit l'agent **"Cohérence Data"**.
  3. Le moteur local extrait les entités numériques et applique des vérifications déterministes (somme des sous-postes = 100%, calcul des taux de variation entre Année N et Année N-1).
  4. Le LLM analyse en parallèle les assertions sémantiques (ex: "forte hausse" alors que le chiffre indique "-2%").
  5. Le système affiche un rapport de diagnostic dans l'overlay classant les alertes par sévérité : *Incohérence certaine (Calcul faux)*, *Incohérence probable (Contresens texte/chiffre)*, *À vérifier*.
  6. L'agent ne modifie pas automatiquement les chiffres mais permet à l'utilisateur de copier le diagnostic.

---

## 14. EXIGENCES FONCTIONNELLES (FR)

### 14.1 Capture & Intégration Hôte

| ID | Description | Priorité (MoSCoW) |
|:---|:---|:---|
| **FR-001** | Le système doit enregistrer un hook global de clavier pour intercepter le raccourci configurable (par défaut `Ctrl+Shift+N`). | **Must** |
| **FR-002** | Le système doit identifier si Microsoft Word est l'application active lors du déclenchement du raccourci. | **Must** |
| **FR-003** | Le Web Add-in doit capturer la sélection au format texte brut et au format HTML structuré via l'API Office.js. | **Must** |
| **FR-004** | En cas de défaillance du Web Add-in, le système doit basculer sur une capture via UI Automation / Presse-papiers. | **Must** |
| **FR-005** | Si aucune sélection n'est active dans Word, l'overlay doit s'ouvrir avec un message guidant l'utilisateur. | **Must** |
| **FR-006** | Le système doit maintenir une communication WebSocket bidirectionnelle persistante entre le Web Add-in et le client desktop sur `ws://127.0.0.1:5284/ws/addin`. | **Must** |

### 14.2 Overlay & Diff Engine

| ID | Description | Priorité (MoSCoW) |
|:---|:---|:---|
| **FR-010** | L'overlay doit s'ouvrir en mode *TopMost* sans voler définitivement le focus contextuel du système. | **Must** |
| **FR-011** | Le moteur de rendu de l'overlay doit calculer un DIFF visuel précis basé sur l'algorithme de Myers (différence mot à mot et caractère par caractère). | **Must** |
| **FR-012** | Le DIFF doit afficher une convention visuelle stricte : rouge barré (suppression), vert souligné (insertion), jaune (changement typographique). | **Must** |
| **FR-013** | L'overlay doit offrir les boutons d'action : `Accepter` (`Enter`), `Rejeter` (`Escape`), `Régénérer`, `Copier dans le presse-papiers`, `Éditer`. | **Must** |
| **FR-014** | L'overlay doit adapter son layout : mode compact contextuel (< 150 mots) ou mode étendu fenêtré (> 150 mots). | **Should** |
| **FR-015** | L'overlay doit être intégralement navigable au clavier (raccourcis Tab, Enter, Echap, Flèches). | **Must** |

### 14.3 Agents Métier Spécialisés

| ID | Description | Priorité (MoSCoW) |
|:---|:---|:---|
| **FR-030** | **Agent Correction :** Doit corriger l'orthographe, la grammaire, la syntaxe, les accords complexes et la ponctuation française/anglaise. | **Must** |
| **FR-031** | **Agent Correction :** Doit intégrer et respecter les règles orthotypographiques professionnelles (espaces insécables devant les doubles ponctuations, majuscules aux institutions, formatage des nombres). | **Must** |
| **FR-032** | **Agent Correction :** Doit appliquer les règles du profil éditorial sélectionné (Académique, Institutionnel, Note Ministérielle, Synthétique). | **Must** |
| **FR-040** | **Agent Reformulation :** Doit proposer 4 modes prédéfinis (Synthétiser -30%, Clarifier/Vulgariser, Style Institutionnel, Restructurer) et un champ d'instruction libre. | **Must** |
| **FR-041** | **Agent Reformulation :** Doit préserver strictement les données factuelles, entités et valeurs chiffrées lors de la réécriture. | **Must** |
| **FR-050** | **Agent Traduction :** Doit traduire bidirectionnellement FR $\leftrightarrow$ EN (variantes sélectionnables : UK et US). | **Must** |
| **FR-051** | **Agent Traduction :** Doit intégrer obligatoirement le glossaire terminologique métier actif injecté dans le contexte. | **Must** |
| **FR-052** | **Agent Traduction :** Doit préserver les entités annotées "Ne pas traduire" (acronymes, noms propres, marques). | **Must** |
| **FR-060** | **Agent Cohérence Data :** Doit extraire les expressions numériques (pourcentages, devises, dates, ordres de grandeur) et exécuter un contrôle déterministe des calculs. | **Must** |
| **FR-061** | **Agent Cohérence Data :** Doit détecter les discordances sémantiques entre le qualificatif textuel et la valeur numérique associée. | **Must** |
| **FR-062** | **Agent Cohérence Data :** Doit formater son résultat sous forme de rapport d'audit textuel non destructif (consultation seule). | **Must** |
| **FR-070** | **Agent Définition / Explication :** Doit générer une fiche explicative synthétique d'un concept, terme technique ou acronyme sélectionné. | **Must** |

### 14.4 Réinjection & Intégrité Documentaire

| ID | Description | Priorité (MoSCoW) |
|:---|:---|:---|
| **FR-080** | Le système doit réinjecter le texte validé dans Word via `Office.js` en utilisant l'instruction `range.insertHtml(html, "Replace")`. | **Must** |
| **FR-081** | Le système doit préserver les attributs typographiques inline de base : gras, italique, souligné, police, taille, couleur, exposant, indice, liens hypertextes. | **Must** |
| **FR-082** | En cas d'échec de la réinjection Office.js, le système doit automatiquement copier le fragment corrigé dans le presse-papiers Windows et notifier l'utilisateur pour un `Ctrl+V` manuel. | **Must** |
| **FR-083** | L'opération de réinjection doit constituer une transaction unitaire dans la pile d'annulation Word afin d'être réversible via un unique `Ctrl+Z`. | **Must** |

---

## 15. EXIGENCES NON-FONCTIONNELLES (NFR)

| ID | Catégorie | Exigence Spécifique | Métrique Cible | MoSCoW |
|:---|:---|:---|:---|:---|
| **NFR-001** | **Ressources** | Empreinte mémoire RAM du processus desktop ALENIA au repos (idle). | $< 150\text{ Mo}$ | **Must** |
| **NFR-002** | **Ressources** | Charge CPU du client desktop en arrière-plan (hors traitement actif). | $< 1.5\%$ | **Must** |
| **NFR-003** | **Performance** | Temps de réponse entre le raccourci clavier et l'affichage complet de l'Overlay. | $< 800\text{ ms}$ | **Must** |
| **NFR-004** | **Performance** | Temps de traitement de la réinjection dans Word après clic sur "Accepter". | $< 400\text{ ms}$ | **Must** |
| **NFR-005** | **Réseau** | Démarrage du serveur Kestrel local et liaison du port d'écoute. | Port fixe `5284` sur `127.0.0.1` | **Must** |
| **NFR-006** | **Affichage** | Support du DPI Scaling multi-écrans sans distorsion visuelle. | 100%, 125%, 150%, 200% | **Must** |
| **NFR-007** | **Thème** | Prise en charge des thèmes Windows Clair et Sombre avec bascule dynamique. | Immédiat sans redémarrage | **Should** |
| **NFR-008** | **Fiabilité** | Reconnexion automatique du WebSocket Web Add-in $\leftrightarrow$ Desktop lors du redémarrage d'un composant. | Retry exponentiel max 3s | **Must** |
| **NFR-009** | **Packaging** | Taille totale de l'installateur MSI tout inclus (.NET 8 Runtime embarqué si requis). | $< 120\text{ Mo}$ | **Should** |

---

## 16. EXIGENCES IA & LLM (AI)

| ID | Description | MoSCoW |
|:---|:---|:---|
| **AI-001** | **Séparation Stricte Données / Instructions :** Le texte documentaire utilisateur doit être encapsulé dans des balises XML fermées (`<document_content>...</document_content>`) avec échappement des caractères spéciaux pour neutraliser les injections de prompt indirectes. | **Must** |
| **AI-002** | **Prompt Hardening :** Les prompts système doivent contenir des clauses de garde interdisant formellement l'exécution d'instructions présentes à l'intérieur du texte utilisateur. | **Must** |
| **AI-003** | **Formatage Structuré des Réponses :** Le LLM doit être contraint de répondre dans un schéma JSON strict ou dans un format balisé normalisé pour permettre un parsing déterministe du texte de remplacement et des métadonnées d'explication. | **Must** |
| **AI-004** | **Détection d'Hallucinations de Sources :** Si l'agent produit une citation, un chiffre non présent dans le texte source ou une URL, le moteur doit ajouter une alerte visuelle : `[Source générée non vérifiée]`. | **Should** |
| **AI-005** | **Modèle de Référence :** Le moteur par défaut doit être configuré sur **Claude 3.5 Sonnet** (Anthropic API). | **Must** |
| **AI-006** | **Mécanisme de Failover :** En cas d'erreur HTTP 5xx, de rate limit 429 ou de timeout (> 20s) sur Claude, la passerelle doit rejouer automatiquement la requête sur **Google Gemini 1.5 Pro** avec adaptation transparente des formats de prompt. | **Must** |
| **AI-007** | **Paramétrage Température :** La température des requêtes doit être fixée dynamiquement par agent : `0.1` pour Correction, `0.2` pour Cohérence Data, `0.3` pour Traduction, `0.7` pour Reformulation. | **Must** |

---

## 17. EXIGENCES DE SÉCURITÉ (SEC)

| ID | Description | MoSCoW |
|:---|:---|:---|
| **SEC-001** | **Stockage des Secrets :** Les clés d'API (Anthropic, Google, BYOK) et la clé de licence doivent être chiffrées localement sur le poste via l'API de protection des données Windows (**DPAPI** - `ProtectedData.Protect` avec portée utilisateur ou machine). | **Must** |
| **SEC-002** | **Isolation Réseau Locale :** Le serveur local Kestrel doit être configuré avec un binding strict sur `IPAddress.Loopback` (`127.0.0.1`). Toute tentative de connexion provenant d'une IP externe doit être rejetée au niveau socket. | **Must** |
| **SEC-003** | **Validation des Requêtes Locales (CORS / Headers) :** Les endpoints HTTP et WebSocket locaux doivent vérifier l'en-tête `Origin` afin de n'autoriser que les requêtes émanant de l'environnement Office Add-in (`https://localhost` / WebView2). | **Must** |
| **SEC-004** | **Chiffrement des Flux Sortants :** Toutes les communications vers les API LLM externes doivent être chiffrées en **TLS 1.3** (ou TLS 1.2 minimum). | **Must** |
| **SEC-005** | **Support Proxy d'Entreprise :** Le client ALENIA doit supporter les proxys HTTP/HTTPS d'entreprise, incluant la détection automatique WPAD, la configuration manuelle (Host, Port) et l'authentification (Basic, NTLM). | **Must** |
| **SEC-006** | **Zero Data Persistence :** Le contenu textuel soumis ne doit jamais être enregistré dans la base SQLite locale ni dans des fichiers temporaires non chiffrés. Seuls les compteurs de tokens, l'horodatage et les identifiants d'actions sont persistés. | **Must** |
| **SEC-007** | **Non-élévation UIPI :** L'application ne doit pas solliciter d'élévation administrateur au runtime afin de garantir une cohabitation fluide et sécurisée avec Word exécuté en contexte utilisateur standard. | **Must** |

---

## 18. EXIGENCES DE CONFIDENTIALITÉ & RGPD (PRIV)

| ID | Description | MoSCoW |
|:---|:---|:---|
| **PRIV-001** | **Minimisation des Données :** Aucune donnée à caractère personnel (DCP) issue des documents Word ne doit transiter vers des serveurs d'analytics. | **Must** |
| **PRIV-002** | **Pseudonymisation des Événements :** L'identifiant utilisateur local est généré par hachage SHA-256 salé non réversible de l'identifiant machine / nom d'utilisateur Windows. | **Must** |
| **PRIV-003** | **Rétention des Données Locales :** Les logs opérationnels et statistiques d'utilisation stockés dans SQLite sont automatiquement purgés selon une politique configurable (par défaut : 90 jours, maximum absolu : 365 jours). | **Must** |
| **PRIV-004** | **Droit à l'Effacement Local :** La console d'administration locale doit comporter un bouton "Purger l'ensemble des données et logs" permettant une remise à zéro complète de la base SQLite. | **Must** |
| **PRIV-005** | **Exclusion de Métriques Intrusives :** Le suivi individuel du temps passé par document ou du contenu rejeté est formellement exclu de l'outil. | **Must** |

---

## 19. MOTEUR DLP & TOKENISATION (DLP)

### 19.1 Pipeline DLP Local Réversible

```
[Texte Brut Word]
       │
       ▼
┌─────────────────────────────────────────────────────────┐
│ 1. MOTEUR D'ANALYSE DLP (Local)                         │
│    - Regex Engines : Emails, Téléphones, Cartes, IBAN   │
│    - Patterns Financiers : Montants (€, $, £, M€, Md€)  │
│    - Dictionnaire / Noms de Projets Métier              │
│    - NER Local (Reconnaissance d'Entités Nommées)       │
└──────────────┬──────────────────────────────────────────┘
               │
               ▼
┌─────────────────────────────────────────────────────────┐
│ 2. TABLE DE MAPPAGE ÉPHÉMÈRE EN RAM                     │
│    {                                                    │
│      "VAL_001": "Jean Dupont",                          │
│      "VAL_002": "45,8 M€",                              │
│      "VAL_003": "Projet HELIOS"                         │
│    }                                                    │
└──────────────┬──────────────────────────────────────────┘
               │
               ▼
[Texte Pseudonymisé] ──► Transmission HTTPS Sécurisée ──► [LLM Cloud]
                                                               │
                                                               ▼
[Texte Modifié avec Placeholders] ◄────────────────────── [Réponse LLM]
       │
       ▼
┌─────────────────────────────────────────────────────────┐
│ 3. MOTEUR DE DÉ-TOKENISATION (Local)                    │
│    Substitution inverse : "VAL_xxx" ──► "Valeur Réelle" │
└──────────────┬──────────────────────────────────────────┘
               │
               ▼
[Texte Restauré Final avec Entités Réelles Préservées] ──► [DIFF Overlay]
```

### 19.2 Exigences Spécifiques DLP

| ID | Description | MoSCoW |
|:---|:---|:---|
| **DLP-001** | Le système doit détecter et pseudonymiser via tokens réversibles les types de données suivants : Adresses e-mail, Numéros de téléphone internationaux, Montants financiers et devises, Noms de projets classifiés (issus d'une liste administrée), Noms de personnes (via NER/Dictionnaires). | **Must** |
| **DLP-002** | Les tokens générés doivent être contextualisés (ex: `<PERSON_1>`, `<AMOUNT_1>`, `<PROJECT_1>`) afin que le LLM conserve le contexte syntaxique, grammatical et sémantique pour accorder correctement les verbes et adjectifs. | **Must** |
| **DLP-003** | La table de correspondance des tokens doit être stockée exclusivement en mémoire vive (RAM) pendant la durée de la session de traitement et détruite immédiatement après la restitution. | **Must** |
| **DLP-004** | L'administrateur doit pouvoir activer ou désactiver chaque catégorie de filtre DLP depuis la console d'administration locale. | **Must** |

---

## 20. MODÈLE DE RÔLES & RBAC

Étant donné le déploiement local pour 10 postes au MVP, la gestion des rôles est structurée localement :

| Rôle | Périmètre d'Accès | Droits Opérationnels |
|:---|:---|:---|
| **Utilisateur (Analyste)** | Client Desktop, Taskpane Word, Overlay | - Exécution de l'ensemble des agents.<br>- Sélection du profil éditorial actif.<br>- Modification des raccourcis clavier personnels.<br>- Ajustement du thème visuel.<br>- Visualisation de son historique de session. |
| **Administrateur Local** | Console Web d'Administration (`/admin`) | - Saisie et modification des clés d'API LLM globales.<br>- Activation de la clé de licence du poste.<br>- Configuration des paramètres Proxy d'entreprise.<br>- Gestion des règles DLP et dictionnaires de mots interdits.<br>- Édition des profils éditoriaux et glossaires d'entreprise.<br>- Consultation des statistiques d'usage et logs d'erreurs.<br>- Exportation et purge des données SQLite. |

L'accès à la console `/admin` est protégé par un mot de passe administrateur défini lors de la première installation du package MSI.

---

## 21. MODULE D'ADMINISTRATION (ADM)

| ID | Description | MoSCoW |
|:---|:---|:---|
| **ADM-001** | Le serveur Kestrel doit servir une interface d'administration Web responsive accessible sur `http://127.0.0.1:5284/admin`. | **Must** |
| **ADM-002** | L'accès à la console d'administration doit être restreint par une authentification par mot de passe avec hachage PBKDF2/BCrypt stocké localement. | **Must** |
| **ADM-003** | La console doit permettre de renseigner, tester et sauvegarder les clés d'API pour **Anthropic** et **Google Gemini**. | **Must** |
| **ADM-004** | La console doit fournir une interface de gestion des **Glossaires** (création de listes de termes bilingues FR/EN avec import/export CSV). | **Must** |
| **ADM-005** | La console doit fournir une interface de gestion des **Profils Éditoriaux** (nom du profil, description, consignes stylistiques injectées au prompt système). | **Must** |
| **ADM-006** | La console doit afficher l'état de la licence logicielle (Identifiant machine, validité, date d'expiration, statut actif/inactif). | **Must** |
| **ADM-007** | La console doit afficher un journal des événements d'erreur système et d'appels API (sans contenu textuel). | **Must** |

---

## 22. MODULE ANALYTICS & CONFIDENTIALITÉ

### 22.1 Métriques Collectées Localement (SQLite)
- `event_id` : Identifiant unique auto-incrémenté.
- `timestamp` : Horodatage UTC de l'action.
- `user_hash` : Hash anonymisé de la machine.
- `agent_type` : Type d'agent sollicité (`CORRECTION`, `REFORMULATION`, `TRADUCTION`, `COHERENCE`, `DEFINITION`).
- `provider_used` : Fournisseur LLM ayant traité la demande (`ANTHROPIC`, `GEMINI`, `LOCAL`).
- `input_word_count` : Nombre de mots du texte source.
- `output_word_count` : Nombre de mots du texte généré.
- `tokens_prompt` : Nombre de tokens consommés en entrée.
- `tokens_completion` : Nombre de tokens consommés en sortie.
- `latency_ms` : Durée totale du traitement en millisecondes.
- `action_result` : Décision utilisateur (`ACCEPTED`, `REJECTED`, `EDITED`, `FAILED`).
- `dlp_matches_count` : Nombre d'entités masquées par le DLP.

---

## 23. MODÈLE DE DONNÉES (DATA MODEL LOCAL)

Base de données locale : **SQLite 3** (`%LOCALAPPDATA%\ALENIA\alenia_data.db`).

```sql
-- Table de configuration globale
CREATE TABLE IF NOT EXISTS system_config (
    config_key TEXT PRIMARY KEY,
    config_value TEXT NOT NULL,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- Table des profils éditoriaux
CREATE TABLE IF NOT EXISTS editorial_profiles (
    profile_id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    description TEXT,
    system_instructions TEXT NOT NULL,
    is_default BOOLEAN DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- Table des glossaires
CREATE TABLE IF NOT EXISTS glossaries (
    entry_id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_term TEXT NOT NULL,
    target_term TEXT NOT NULL,
    domain TEXT DEFAULT 'GENERAL',
    case_sensitive BOOLEAN DEFAULT 1,
    do_not_translate BOOLEAN DEFAULT 0
);

-- Table des règles DLP
CREATE TABLE IF NOT EXISTS dlp_rules (
    rule_id TEXT PRIMARY KEY,
    rule_name TEXT NOT NULL,
    category TEXT NOT NULL, -- 'REGEX', 'KEYWORD', 'NER'
    pattern_value TEXT NOT NULL,
    replacement_tag TEXT NOT NULL,
    is_enabled BOOLEAN DEFAULT 1
);

-- Table des métriques d'usage (Analytics anonymisées)
CREATE TABLE IF NOT EXISTS usage_analytics (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
    agent_id TEXT NOT NULL,
    provider TEXT NOT NULL,
    input_words INTEGER NOT NULL,
    output_words INTEGER NOT NULL,
    prompt_tokens INTEGER,
    completion_tokens INTEGER,
    latency_ms INTEGER NOT NULL,
    user_action TEXT NOT NULL -- 'ACCEPTED', 'REJECTED', 'EDITED'
);

-- Table du journal d'audit et erreurs
CREATE TABLE IF NOT EXISTS system_logs (
    log_id INTEGER PRIMARY KEY AUTOINCREMENT,
    log_timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
    log_level TEXT NOT NULL, -- 'INFO', 'WARN', 'ERROR', 'FATAL'
    component TEXT NOT NULL,
    message TEXT NOT NULL,
    exception_details TEXT
);
```

---

## 24. ARCHITECTURE DES AGENTS

Chaque agent implémente une interface standard .NET :

```csharp
public interface IAleniaAgent
{
    string AgentId { get; }
    string DisplayName { get; }
    Task<AgentExecutionResult> ExecuteAsync(AgentContext context, CancellationToken ct);
}
```

### 24.1 Matrice de Paramétrage des Agents

| Agent | Prompt Système Spécifique | Température | Sortie Formatée | Action Document |
|:---|:---|:---:|:---:|:---:|
| **Correction** | Rigueur orthotypographique, préservation stricte du sens et du ton. Respect des espaces insécables et majuscules institutionnelles. | `0.1` | HTML avec balises de modification | Réinjection Word |
| **Reformulation**| Réécriture stylistique selon consigne (concision, clarté, registre). Interdiction d'omettre des faits ou chiffres. | `0.7` | Texte HTML remanié | Réinjection Word |
| **Traduction** | Traduction fidèle respectant la terminologie du glossaire injecté. Adaptation culturelle des conventions numériques (ex: 1,000.50 vs 1 000,50). | `0.3` | Texte HTML traduit | Réinjection Word |
| **Cohérence Data** | Analyse critique des calculs, pourcentages et assertions factuelles. Extraction et vérification arithmétique. | `0.2` | Rapport Markdown structuré | Consultation Overlay |
| **Définition** | Synthèse conceptuelle, étymologie, contexte géopolitique/économique et usages recommandés. | `0.4` | Fiche synthétique balisée | Consultation Overlay |

---

## 25. LLM GATEWAY & MÉCANISME DE FALLBACK

```
[Requête Agent Préparée]
           │
           ▼
┌────────────────────────────────────────────────────────┐
│ LLM GATEWAY (.NET Core HttpClient Engine)               │
│ - Injection Proxy Corporate                            │
│ - Timeout HttpClient : 20 000 ms                       │
└──────────┬─────────────────────────────────────────────┘
           │
           ├─────────────────────────────────────────────┐
           │ (Tentative 1 : Fournisseur Primaire)         │
           ▼                                             │
┌──────────────────────────────────────┐                 │
│ Anthropic API Client                 │                 │
│ Endpoint: api.anthropic.com/v1/messages               │
│ Modèle: claude-3-5-sonnet-latest     │                 │
└──────────────────┬───────────────────┘                 │
                   │                                     │
       ┌───────────┴───────────┐                         │
       ▼                       ▼                         │
 [Réponse 200 OK]        [Erreur / Timeout / 429]        │
       │                       │                         │
       │                       ▼                         │
       │         ┌──────────────────────────────────────┐│
       │         │ (Tentative 2 : Bascule Automatique)  ││
       │         │ Google Gemini Client                 ││
       │         │ Endpoint: generativelanguage...      ││
       │         │ Modèle: gemini-1.5-pro-latest        ││
       │         └──────────────────┬───────────────────┘│
       │                            │                    │
       │                ┌───────────┴───────────┐        │
       │                ▼                       ▼        │
       │          [Réponse 200 OK]        [Échec Total]  │
       │                │                       │        │
       ▼                ▼                       ▼        │
┌───────────────────────────────┐ ┌─────────────────────┴┐
│ Traitement Normal & Affichage │ │ Notification Erreur  │
│ du DIFF dans l'Overlay        │ │ "Service Indisponible│
└───────────────────────────────┘ └──────────────────────┘
```

---

## 26. INTÉGRATION MICROSOFT WORD (WEB ADD-IN)

### 26.1 Spécification Technique de l'Add-in
- **Type :** Office JavaScript Web Add-in (compatible Word Desktop Windows et M365).
- **Fichier de Déploiement :** `manifest.xml` configuré avec les permissions `ReadWriteDocument`.
- **Mécanisme d'Exécution :** Exécuté dans le runtime WebView2 intégré de Microsoft Word.
- **Hébergement des Assets Add-in :** Les fichiers HTML, JS (`taskpane.html`, `taskpane.js`, `office.js`) sont servis directement par le serveur Kestrel local sur `http://127.0.0.1:5284/addin/`.

### 26.2 Séquence JavaScript Office.js

```javascript
// Lecture de la sélection avec balisage HTML
async function captureWordSelection() {
    return new Promise((resolve, reject) => {
        Word.run(async (context) => {
            const range = context.document.getSelection();
            const htmlResult = range.getHtml();
            const textResult = range.getText();
            await context.sync();
            resolve({
                text: textResult.value,
                html: htmlResult.value
            });
        }).catch(reject);
    });
}

// Réinjection formatée dans la sélection active
async function reinjectWordSelection(cleanHtml) {
    return Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertHtml(cleanHtml, Word.InsertLocation.replace);
        await context.sync();
    });
}
```

---

## 27. MATRICE DE COMPATIBILITÉ

| Environnement / Composant | Version Supportée | Niveau de Support | Observations |
|:---|:---|:---:|:---|
| **Windows OS** | Windows 10 (21H2+), Windows 11 | **Complet** | x64 uniquement. |
| **Microsoft Word Desktop** | Microsoft 365 Apps, Office LTSC 2021 | **Complet** | Intégration Taskpane + Sideload Manifest. |
| **Microsoft Word Online** | Navigateurs Edge, Chrome | **Partiel** | Taskpane opérationnelle, pas d'overlay WPF natif. |
| **PowerPoint / Outlook** | M365 / LTSC 2021 | **V1.5** | Hors périmètre MVP. |
| **macOS (Word pour Mac)** | macOS 13+ (Ventura, Sonoma) | **V2** | Client Desktop à porter, Add-in déjà compatible. |
| **Proxy Réseau** | Proxy standard, Authentifié NTLM/Basic | **Complet** | Paramétrable dans la console admin. |
| **Résolutions Écran** | Full HD (1080p), 2K, 4K multi-écrans | **Complet** | Gestion per-monitor DPI scaling. |

---

## 28. SPÉCIFICATIONS UX/UI

### 28.1 Principes Fondamentaux d'Interface
1. **Légèreté & Discrétion :** L'interface ne doit pas masquer la zone de travail Word en cours d'édition.
2. **Contraste Visuel Élevé :** Lisibilité parfaite des différences textuelles selon les standards WCAG 2.1 AA.
3. **Clavier First :** Chaque action est associée à un raccourci mnémotechnique sans nécessiter la souris.

---

## 29. WIREFRAMES TEXTUELS COMPLETS

### 29.1 Taskpane Microsoft Word (Panneau Latéral Droit)

```
┌──────────────────────────────────────────────┐
│  ALENIA — Assistant Rédactionnel     [─] [✕] │
├──────────────────────────────────────────────┤
│  ● Connecté (127.0.0.1:5284)                 │
├──────────────────────────────────────────────┤
│  SÉLECTION ACTIVE :                          │
│  ┌────────────────────────────────────────┐  │
│  │ "Le taux d'inflation sous-jacent a     │  │
│  │ progressé de 2,4% au T3, confirmant... │  │
│  │ (3 paragraphes - 142 mots)             │  │
│  └────────────────────────────────────────┘  │
│                                              │
│  ACTIONS DISPONIBLES :                       │
│  ┌──────────────────┐  ┌──────────────────┐  │
│  │ [✏️ Corriger]     │  │ [🔄 Reformuler]  │  │
│  └──────────────────┘  └──────────────────┘  │
│  ┌──────────────────┐  ┌──────────────────┐  │
│  │ [🌐 Traduire ▾]  │  │ [📊 Cohérence]   │  │
│  └──────────────────┘  └──────────────────┘  │
│  ┌────────────────────────────────────────┐  │
│  │ [📖 Définir / Expliquer le concept]    │  │
│  └────────────────────────────────────────┘  │
│                                              │
│  PROFIL ÉDITORIAL ACTIF :                    │
│  [ Note Ministérielle / Synthèse      ▾ ]    │
│                                              │
├──────────────────────────────────────────────┤
│  HISTORIQUE RÉCENT DE SESSION :              │
│  • 11:42 - Correction (142 mots)  [Accepté]  │
│  • 11:35 - Traduction EN-UK       [Accepté]  │
│  • 11:15 - Reformulation          [Rejeté]   │
├──────────────────────────────────────────────┤
│  [⚙️ Admin Locale]       [Raccourci: Ctrl+⇧+N]│
└──────────────────────────────────────────────┘
```

### 29.2 Fenêtre Overlay WPF (Visualisation DIFF & Validation)

```
┌──────────────────────────────────────────────────────────────────────────────┐
│  ALENIA — Proposition de Modification (Agent : Correction)           [─] [✕] │
├──────────────────────────────────────────────────────────────────────────────┤
│  Profil : Institutionnel | Fournisseur : Claude 3.5 | Temps : 1.2s           │
├──────────────────────────────────────────────────────────────────────────────┤
│  DIFF VISUEL (Comparaison de la sélection) :                                 │
│  ┌────────────────────────────────────────────────────────────────────────┐  │
│  │ Le taux d'inflation ~~sous jacent~~ **sous-jacent** a progressé de     │  │
│  │ ~~2,4%~~ **2,4 %** au ~~troisieme~~ **troisième** trimestre, ce qui    │  │
│  │ ~~montre~~ **témoigne d'** une stabilisation de l'indice des prix.     │  │
│  │                                                                        │  │
│  │ Légende : ~~Suppression~~ (Rouge) | **Ajout** (Vert) | Détection DLP : 0│  │
│  └────────────────────────────────────────────────────────────────────────┘  │
│                                                                              │
│  Zone d'Édition Manuelle (Optionnelle) :                                     │
│  ┌────────────────────────────────────────────────────────────────────────┐  │
│  │ Le taux d'inflation sous-jacent a progressé de 2,4 % au troisième...   │  │
│  └────────────────────────────────────────────────────────────────────────┘  │
│                                                                              │
│  Consigne d'Ajustement / Régénération :                                      │
│  [ Rendre la conclusion plus percutante...                       ] [🔄 Relancer]│
├──────────────────────────────────────────────────────────────────────────────┤
│  [  ✅ ACCEPTER (Entrée)  ]   [  ❌ REJETER (Échap)  ]   [  📋 Copier Texte ] │
└──────────────────────────────────────────────────────────────────────────────┘
```

---

## 30. GESTION DES ERREURS (ERROR HANDLING)

| Code d'Erreur | Cause Identifiée | Comportement Système | Message Utilisateur Affiché |
|:---|:---|:---|:---|
| **ERR_NO_SELECTION** | L'utilisateur déclenche le raccourci sans sélection active dans Word. | L'overlay s'ouvre en mode guidage sans solliciter le réseau. | "Veuillez sélectionner le texte à analyser dans Word avant de lancer ALENIA." |
| **ERR_WS_DISCONNECTED** | Le Web Add-in n'arrive pas à joindre le serveur Kestrel local. | Tentative de reconnexion automatique en boucle (3x). | "Connexion au moteur ALENIA interrompue. Vérifiez que l'application de bureau est bien active." |
| **ERR_LLM_TIMEOUT** | Absence de réponse du LLM primaire après 20 secondes. | Déclenchement du fallback immédiat vers Google Gemini. | Notification discrète : "Bascule sur le fournisseur de secours en cours..." |
| **ERR_ALL_PROVIDERS_DOWN**| Échec combiné Anthropic et Google Gemini (réseau coupé). | Interruption du pipeline, libération de la mémoire. | "Impossible de joindre les services d'IA. Vérifiez votre connexion Internet ou la configuration Proxy." |
| **ERR_REINJECT_FAILED** | Conflit Word lors de l'exécution de `insertHtml`. | Copie automatique du résultat dans le presse-papiers Windows. | "La réinjection automatique a échoué. Le texte corrigé a été copié dans votre presse-papiers (Ctrl+V)." |
| **ERR_PORT_OCCUPIED** | Le port `5284` est déjà réservé par un autre processus au lancement. | Notification système Tray, blocage propre du démarrage. | "Le port local 5284 est occupé. Veuillez libérer ce port ou modifier la configuration." |

---

## 31. MODES HORS-LIGNE & DÉGRADÉ

1. **Absence de Connexion Internet :**
   - Le moteur DLP et le serveur Kestrel démarrent normalement.
   - Les agents LLM distants sont marqués "Indisponibles".
   - L'agent **Cohérence Data (Moteur Déterministe)** reste actif pour les vérifications arithmétiques pures locales.
   - L'utilisateur est clairement notifié de l'impossibilité de solliciter la génération cloud.
2. **Mode Dégradé Presse-papiers (Clipboard Fallback) :**
   - Si l'environnement Word désactive l'Add-in (ex: blocage temporaire de script), ALENIA fonctionne en mode autonome via `Ctrl+C` / `Ctrl+V` standard avec l'Overlay.

---

## 32. EXIGENCES DE PERFORMANCE (PERF)

| ID | Indicateur Clé | Condition de Test | Seuil Limite Toléré |
|:---|:---|:---|:---|
| **PERF-001** | Temps de démarrage à froid | Lancement au démarrage de Windows jusqu'à l'icône Tray active. | $< 2.5\text{ s}$ |
| **PERF-002** | Latence d'ouverture Overlay | Après appui physique sur `Ctrl+Shift+N`. | $< 800\text{ ms}$ |
| **PERF-003** | Latence locale du moteur DLP | Analyse Regex + Tokenisation sur 1 000 mots. | $< 50\text{ ms}$ |
| **PERF-004** | Durée de calcul du DIFF | Calcul de différence Myers sur 1 000 mots. | $< 100\text{ ms}$ |
| **PERF-005** | Consommation RAM au repos | 4 heures d'inactivité en arrière-plan. | $< 150\text{ Mo}$ |
| **PERF-006** | Fuite mémoire (Memory Leak) | Exécution en boucle de 100 requêtes consécutives. | Croissance résiduelle $< 5\text{ Mo}$ |

---

## 33. OBSERVABILITÉ & JOURNALISATION (OBS)

- **Fichier de Journalisation Local :** `%LOCALAPPDATA%\ALENIA\logs\alenia_app.log` (Rotation quotidienne, taille maximale 10 Mo, conservation 7 jours).
- **Format des Logs :** JSON structuré (Serilog).
- **Niveaux de Gravité :** `DEBUG`, `INFO`, `WARNING`, `ERROR`, `FATAL`.
- **Règle Absolue d'Audit :** Aucun texte utilisateur, aucun nom de document, aucun token dé-tokenisé ne doit apparaître dans les lignes de log.
- **Exemple de Log Conforme :**
```json
{
  "timestamp": "2023-10-24T14:32:01.104Z",
  "level": "INFO",
  "component": "LLMGateway",
  "action": "ExecuteRequest",
  "provider": "Anthropic",
  "model": "claude-3-5-sonnet",
  "prompt_tokens": 420,
  "completion_tokens": 185,
  "latency_ms": 1142,
  "status_code": 200
}
```

---

## 34. THREAT MODEL (MODÈLE DE MENACES)

| ID | Vecteur de Menace | Risque Identifié | Impact | Mesure de Sécurité Implémentée |
|:---|:---|:---|:---:|:---|
| **TM-01** | **Fuite de Données Confidentielles (Data Exfiltration)** | Transmission de données C2 non masquées au fournisseur cloud. | Élevé | Pipeline DLP local obligatoire avec tokenisation réversible avant émission réseau. |
| **TM-02** | **Indirect Prompt Injection** | Présence d'instructions malveillantes cachées dans le texte Word analysé. | Élevé | Encapsulation stricte du texte dans des balises XML fermées et instructions système hermétiques. |
| **TM-03** | **Attaque Cross-Site (CSRF Local)** | Un script malveillant dans un navigateur tente d'appeler `localhost:5284`. | Élevé | Binding `127.0.0.1`, vérification des en-têtes `Origin` et validation de token de session local. |
| **TM-04** | **Vol des Clés d'API LLM** | Lecture non autorisée du fichier de configuration sur le disque. | Critique | Chiffrement systématique de toutes les clés d'API via **Windows DPAPI**. |
| **TM-05** | **Déni de Service Local (Port Hijacking)** | Une application tierce écoute sur le port `5284` pour intercepter les flux. | Moyen | Utilisation de socket exclusif (`SO_EXCLUSIVEADDRUSE`) sous Windows. |

---

## 35. CAS LIMITES (EDGE CASES)

1. **Sélection Massive (> 5 000 mots) :** Le système affiche un avertissement préventif indiquant que le traitement par morceaux peut altérer la cohérence globale et suggère de procéder section par section.
2. **Sélection Exclusivement Numérique ou Tableau :** L'agent **Cohérence Data** est automatiquement activé par défaut.
3. **Texte Contenant des Caractères de Contrôle ou Émoticônes :** Normalisation UTF-8 stricte avant encodage JSON.
4. **Document Word Verrouillé en Lecture Seule :** L'action de réinjection est désactivée ; seul le bouton "Copier dans le presse-papiers" est proposé.
5. **Multiples Instances de Word Ouvertes :** La Taskpane communique son identifiant d'instance unique lors du handshake WebSocket afin que l'overlay réinjecte exactement dans le document émetteur.

---

## 36. USER STORIES (US)

### US-001 : Correction Orthographique & Stylistique
**En tant qu'** Analyste Rédacteur,  
**Je veux** corriger les fautes d'orthographe, de grammaire et de typographie dans mon paragraphe Word sélectionné en un seul raccourci,  
**Afin de** fiabiliser immédiatement la qualité de ma note sans casser la mise en forme de mon document.

### US-002 : Reformulation selon une Consigne Libre
**En tant qu'** Économiste,  
**Je veux** demander la reformulation d'un paragraphe technique avec une consigne spécifique (ex: "Rendre accessible à un ministre"),  
**Afin d'** adapter mon niveau de discours en conservant l'intégralité des données factuelles.

### US-003 : Traduction Terminologique Spécialisée
**En tant qu'** Analyste Bilingue,  
**Je veux** traduire mes conclusions en anglais britannique en appliquant le glossaire institutionnel officiel,  
**Afin de** garantir l'exactitude des termes économiques sans erreurs d'équivalence.

### US-004 : Détection d'Erreurs de Données Chiffrées
**En tant qu'** Expert Chiffres,  
**Je veux** soumettre un tableau textuel pour vérifier que les variations de pourcentages et les sommes sont cohérentes,  
**Afin d'** éviter de publier des données arithmétiquement contradictoires.

### US-005 : Configuration Simple et Sécurisée par l'Administrateur
**En tant qu'** Administrateur Système,  
**Je veux** renseigner les clés API d'entreprise et les règles DLP sur une console locale sécurisée,  
**Afin de** déployer l'outil sans nécessiter de serveur cloud externe dédié.

---

## 37. CRITÈRES D'ACCEPTATION (GHERKIN)

### CA-01 : Validation du DIFF et Réinjection dans Word

```gherkin
SCÉNARIO: Correction acceptée avec succès dans Microsoft Word
  GIVEN Microsoft Word est ouvert avec un document contenant "L'inflation a progressé de 2,4%."
  AND Le mot "progressé" est sélectionné
  AND Le client ALENIA Desktop est actif sur le port local 5284
  WHEN L'utilisateur appuie sur la combinaison de touches "Ctrl+Shift+N"
  THEN La fenêtre Overlay ALENIA s'affiche en moins de 800 ms
  AND L'Overlay présente une proposition de modification
  WHEN L'utilisateur appuie sur la touche "Entrée" (Accepter)
  THEN La fenêtre Overlay se ferme
  AND Le texte dans Microsoft Word est mis à jour instantanément
  AND La mise en forme (police, taille, couleur) d'origine est rigoureusement conservée
  AND L'action est annulable dans Word par la combinaison "Ctrl+Z"
```

### CA-02 : Protection des Données Sensibles (DLP)

```gherkin
SCÉNARIO: Anonymisation réversible d'un montant confidentiel et d'un nom de projet
  GIVEN Une sélection Word contenant "Le budget alloué au Projet HELIOS est de 15,4 M€."
  AND La règle DLP "Projets Confidentiels" et "Montants Financiers" est active
  WHEN L'utilisateur déclenche l'agent "Reformuler"
  THEN La charge utile transmise au LLM distant contient "<PROJECT_1>" à la place de "Projet HELIOS"
  AND La charge utile transmise contient "<AMOUNT_1>" à la place de "15,4 M€"
  AND Le texte reconstitué dans l'Overlay ALENIA affiche la phrase reformulée contenant explicitement "Projet HELIOS" et "15,4 M€"
```

---

## 38. STRATÉGIE DE TEST GLOBALE

Compte tenu de la contrainte selon laquelle les développements initiaux et les tests de recette sont opérés sur une **machine distante (VM Windows 10/11 dédiée)**, la stratégie s'articule autour de 5 niveaux de validation automatisables et documentés :

1. **Tests Unitaires (.NET 8 - xUnit / FluentAssertions) :** Couverture minimale de 80% sur les moteurs DLP, calcul du DIFF, parsing des prompts et sérialisation JSON.
2. **Tests d'Intégration API (TestServer Kestrel) :** Validation des endpoints `/api/*`, de la liaison WebSocket SignalR et des mécanismes de failover HTTP vers les LLMs mockés.
3. **Tests UI Automation (FlaUI / Appium Windows Driver) :** Automatisation du déclenchement des raccourcis globaux, de la capture de focus et des clics de validation sur l'overlay WPF.
4. **Tests de Sécurité & Non-Régression Prompt Injection :** Exécution d'un banc de test automatisé de 50 attaques d'injection de prompt pour valider l'étanchéité des délimiteurs XML.
5. **Tests de Performance & Mémoire (dotMemory / PerfMon) :** Mesure continue de la mémoire RAM sur 50 cycles de traitement pour garantir le plafond de 150 Mo.

---

## 39. STRATÉGIE UAT & CAHIER DE RECETTE DÉTAILLÉ

### 39.1 Corpus Documentaire de Référence (Fourni aux Testeurs)
Un jeu de **20 documents Word étalons** (.docx) représentatifs du métier d'analyste économique et géopolitique est mis à disposition sur l'environnement de test :
- 5 Notes de conjoncture avec tableaux et pourcentages.
- 5 Analyses prospectives géopolitiques avec acronymes rares.
- 5 Notes de synthèse ministérielles soumises à un style protocolaire strict.
- 5 Documents "pièges" contenant des failles de ponctuation, des calculs volontairement faux et des injections de prompt.

### 39.2 Fiches de Tests d'Acceptation Utilisateur (UAT)

#### TEST-UAT-01 : Installation et Initialisation du Pilote
- **Objectif :** Valider la procédure d'installation MSI et l'initialisation du serveur local.
- **Prérequis :** VM Windows 10 propre sans installation préalable d'ALENIA.
- **Étapes :**
  1. Exécuter `ALENIA-Setup-1.0.msi` en tant qu'administrateur.
  2. Vérifier la présence de l'icône dans la zone de notification (System Tray).
  3. Ouvrir le navigateur et accéder à `http://127.0.0.1:5284/admin`.
  4. Saisir la clé de licence de test à 10 postes.
  5. Saisir une clé API Anthropic valide.
- **Résultat Attendu :** Statut "Licence Active", icône Tray verte "ALENIA Connecté".

#### TEST-UAT-02 : Sideloading du Manifest et Handshake Word
- **Objectif :** Valider l'intégration du Web Add-in dans Word.
- **Prérequis :** Word 365 ou LTSC 2021 lancé.
- **Étapes :**
  1. Aller dans `Accueil` > `Compléments` > `Mes compléments` > `Charger un complément`.
  2. Sélectionner `C:\Program Files\ALENIA\addin\manifest.xml`.
  3. Vérifier l'ouverture de la Taskpane ALENIA à droite.
- **Résultat Attendu :** La Taskpane affiche l'indicateur "● Connecté (127.0.0.1:5284)".

#### TEST-UAT-03 : Correction avec Préservation de Mise en Forme Enrichie
- **Objectif :** Valider que la réinjection via `insertHtml` ne détruit pas le formatage.
- **Prérequis :** Document de test ouvert contenant du texte avec **Gras**, *Italique*, texte en couleur Rouge et un lien hypertexte.
- **Étapes :**
  1. Sélectionner l'ensemble du paragraphe formaté comportant une faute d'orthographe.
  2. Appuyer sur `Ctrl+Shift+N`.
  3. Valider la correction proposée dans l'overlay via la touche `Entrée`.
- **Résultat Attendu :** La faute est corrigée dans Word. Le gras, l'italique, la couleur rouge et le lien hypertexte sont parfaitement conservés à l'identique.

#### TEST-UAT-04 : Contrôle du Failover Transparent (Claude $\rightarrow$ Gemini)
- **Objectif :** Vérifier la résilience du système en cas de coupure de l'API primaire.
- **Prérequis :** Clés Anthropic et Google renseignées dans l'administration.
- **Étapes :**
  1. Simuler une coupure de l'endpoint Anthropic (ex: clé invalide temporaire).
  2. Sélectionner un paragraphe et lancer l'agent "Traduire".
- **Résultat Attendu :** L'overlay affiche le résultat traduit en moins de 4 secondes sans message de crash bloquant. Le log local atteste du basculement sur Google Gemini.

---

## 40. DÉFINITION DU MVP

Le **MVP d'ALENIA** est strictement circonscrit aux éléments suivants :
1. Client Desktop Windows (.NET 8 WPF) intégrant le serveur Kestrel sur `127.0.0.1:5284`.
2. Web Add-in Word (Taskpane latérale + Sideload Manifest).
3. Overlay WPF avec calcul de DIFF visuel dynamique.
4. Les 5 Agents fondamentaux : Correction, Reformulation, Traduction FR/EN, Cohérence Data (déterministe), Définition.
5. Passerelle d'appel direct vers Anthropic Claude avec repli sur Google Gemini.
6. Moteur DLP local par Regex et Tokenisation.
7. Console d'administration Web locale sur le poste.
8. Packaging MSI tout-en-un configuré pour un déploiement pilote de 10 postes.

---

## 41. ROADMAP PRODUIT (MVP, V1.5, V2)

```
2023 - Q4 : MVP (PILOTE 10 POSTES)
├── Client Desktop .NET 8 / Overlay WPF
├── Intégration Word (Web Add-in Taskpane)
├── 5 Agents Fondamentaux (Correction, Reformul., Trad., Cohérence, Définition)
├── DLP Local Réversible (Regex/NER)
├── Passerelle Directe Cloud (Claude 3.5 + Fallback Gemini)
└── Console Admin Locale & Déploiement Sideload

2024 - Q2 : VERSION 1.5 (INDUSTRIALISATION OFFICE)
├── Mode "ALL" (Traitement séquentiel de documents complets)
├── Support Multi-Applications : PowerPoint, Outlook Desktop & Notepad
├── Support des Modèles Locaux (Ollama / Llama 3 On-Premise)
├── Déploiement Centralisé de l'Add-in (Catalogue M365 Admin Center)
├── Intégration Annuaire Entreprise (SSO Azure AD / Entra ID)
└── Détection DLP avancée (Données Médicales & Formats Bancaires)

2024 - Q4 : VERSION 2.0 (SOUVERAINETÉ & CROSS-PLATFORM)
├── Portage macOS (Client Natif Mac + Word Mac Add-in)
├── Backend Centralisé Cloud Souverain Européen (SecNumCloud / HDS)
├── Moteur Vision / OCR (Analyse contextuelle des graphiques et schémas)
├── Module Avancé de Détection de Biais Cognitifs et Rhétoriques
└── Connecteurs DLP d'Entreprise (Microsoft Purview / Symantec)
```

---

## 42. REGISTRE DES RISQUES & PLANS DE MITIGATION

| ID | Intitulé du Risque | Prob. | Gravité | Score | Plan d'Atténuation / Mesure de Contingence |
|:---|:---|:---:|:---:|:---:|:---|
| **RSK-01** | **Altération de formats Word complexes** (tableaux imbriqués, styles rares) lors du `insertHtml`. | Moy. | Élevée | **Élevé** | Détection automatique de la complexité du DOM. Si risque d'altération détecté, notification à l'utilisateur et bascule sur le mode copie sécurisée dans le presse-papiers. |
| **RSK-02** | **Conflit ou blocage du port `5284`** par un logiciel de sécurité / pare-feu local. | Faible | Élevée | **Moyen** | Utilisation de l'API `HttpListener` avec liaison exclusive sur `127.0.0.1`. Documentation de la règle d'exception locale dans le guide d'installation MSI. |
| **RSK-03** | **Refus des flux LLMs sortants** par le Proxy Corporate d'entreprise. | Moy. | Critique | **Élevé** | Support natif des protocoles d'authentification proxy d'entreprise (NTLM/Kerberos/Basic) configurable dans la console d'administration. |
| **RSK-04** | **Rejet utilisateur lié à des hallucinations de données chiffrées**. | Moy. | Élevée | **Élevé** | Moteur déterministe obligatoire en amont du LLM pour l'agent Cohérence Data ; interdiction de modifier les chiffres sans alerte explicite. |
| **RSK-05** | **Latence réseau dégradée** lors des heures de pointe sur les API Cloud. | Moy. | Moy. | **Moyen** | Timeout strict à 20s et mécanisme de bascule automatique et transparente vers l'API Gemini. |

---

## 43. ARBRE DES DÉPENDANCES

```
[Package MSI ALENIA]
  ├── Prérequis OS : Windows 10/11 x64
  ├── Runtime : .NET 8 Desktop Runtime (embarqué)
  ├── WebView2 Runtime (Microsoft Edge)
  └── Base Locale : SQLite 3 Engine
        │
        ▼
[Service ALENIA Core Desktop] ◄── Intègre ──► [Serveur Kestrel (Port 5284)]
  ├── Win32 Global Hooks                       ├── WebSocket Hub SignalR
  ├── Moteur WPF Overlay                       ├── REST API Endpoints (/api)
  ├── Moteur DLP Local                         └── Static Files Web Server (/addin)
  └── Cryptographie DPAPI                                │
        │                                                │
        │ Communication Locale                           │ Chargement Manifest
        ▼                                                ▼
[Overlay WPF Natif]                             [Microsoft Word Desktop]
                                                  └── [JS Web Add-in Taskpane]
                                                            │
                                                            ▼
                                                [API Office.js Runtime]
```

---

## 44. REGISTRE DES QUESTIONS OUVERTES & RÉSOLUTIONS

Toutes les questions ouvertes identifiées lors des cycles de cadrage précédents ont été définitivement tranchées et validées pour le MVP :

| Réf. Question | Sujet Traité | Décision Finale Validée |
|:---|:---|:---|
| **OQ-001** | Normalisation de la marque logicielle | **ALENIA** (orthographe unique validée). |
| **OQ-002** | Technologie d'intégration Word | **JS Web Add-in (Office.js)** privilégié pour l'évolutivité. |
| **OQ-003** | Hébergement du Backend MVP | **100% Local** via Kestrel in-process sur le poste. |
| **OQ-004** | Hébergement Cloud Souverain | Reporté en **Version 2.0**. |
| **OQ-005** | Volume de licences pilotes | Dimensionnement pour **10 postes utilisateurs**. |
| **OQ-006** | Raccourci d'activation par défaut | **`Ctrl+Shift+N`** (avec détection de conflit locale). |
| **OQ-007** | Gestion du mode hors-ligne | **Message d'erreur explicite sans file d'attente**. |
| **OQ-008** | Présence visuelle dans Word | **Taskpane latérale permanente** + Overlay dynamique. |
| **OQ-009** | Configuration du port local | **Port fixe `5284`** lié sur `127.0.0.1`. |
| **OQ-010** | Mode de déploiement de l'Add-in | **Sideloading manuel du fichier `manifest.xml`**. |

---

## 45. DECISION LOG CONSOLIDÉ (DEC-001 À DEC-025)

| ID | Domaine | Décision Validée | Justification Technique & Métier | Statut |
|:---|:---|:---|:---|:---:|
| **DEC-001** | Format | Préservation format de base (MUST) / complexe (SHOULD). | Limitations intrinsèques d'Office.js `insertHtml`. | **Decided** |
| **DEC-002** | Intégration | Approche Hybride (Taskpane Add-in + Overlay WPF). | Meilleur compromis visibilité métier et rapidité d'action. | **Decided** |
| **DEC-003** | Mode ALL | Reporté en Version 1.5. | Priorité absolue à la stabilité sur les sélections courtes. | **Decided** |
| **DEC-004** | Architecture | Client Lourd avec Backend Local Kestrel embarqué. | Zéro coût d'infrastructure serveur cloud pour le pilote. | **Decided** |
| **DEC-005** | Cible | Analystes Grandes Entreprises / Secteur Public. | Exigence stricte de sécurité des données (C2). | **Decided** |
| **DEC-006** | Modèle LLM | Endpoint ALENIA par défaut (Claude 3.5) + Option BYOK. | Simplicité d'onboarding immédiat des 10 testeurs. | **Decided** |
| **DEC-007** | Hôtes | Microsoft Word Desktop uniquement au MVP. | Focalisation sur 85% du volume de rédaction d'analyses. | **Decided** |
| **DEC-008** | Workflow ALL | Analyse séquentielle confirmée pour la V1.5. | Évite la corruption du document complet. | **Decided** |
| **DEC-009** | Réseau | Prise en charge obligatoire des Proxys d'entreprise. | Environnements d'entreprise hautement sécurisés. | **Decided** |
| **DEC-010** | Marque | Nom officiel fixé à "ALENIA". | Cohérence de communication et de documentation. | **Decided** |
| **DEC-011** | Mobilité | Abandon définitif du support Android. | Hors cible métier (rédaction de rapports complexes). | **Decided** |
| **DEC-012** | Privacy | Exclusion du tracking du temps passé par document. | Respect strict de la vie privée des analystes. | **Decided** |
| **DEC-013** | Stack Tech | C# .NET 8 / WPF (Client) + JS Office.js (Add-in). | Performance, robustesse et interopérabilité Windows. | **Decided** |
| **DEC-014** | Souveraineté | Cloud souverain européen planifié pour la V2. | Inutile pour la phase pilote 10 postes. | **Decided** |
| **DEC-015** | Failover | Google Gemini 1.5 Pro retenu comme solution de repli. | Excellente performance bilingue et haute disponibilité. | **Decided** |
| **DEC-016** | Add-in Type | Web Add-in (HTML/JS) plutôt que COM VSTO. | Facilite le portage futur vers macOS en V2. | **Decided** |
| **DEC-017** | Backend Host | Exécution in-process sur la machine locale. | Élimine les contraintes RGPD de transfert tiers. | **Decided** |
| **DEC-018** | Localisation | Hébergement souverain européen ciblé post-MVP. | Conformité réglementaire future. | **Decided** |
| **DEC-019** | Licences | Gestion locale d'un pool de 10 postes par Hardware ID. | Adapté au périmètre contractuel du pilote. | **Decided** |
| **DEC-020** | Raccourci | `Ctrl+Shift+N` fixé par défaut (configurable). | Standard ergonomique non conflictuel dans Word. | **Decided** |
| **DEC-021** | Mode Offline | Erreur explicite immédiate sans file d'attente. | Transparence totale pour l'utilisateur. | **Decided** |
| **DEC-022** | Rendu Diff | Moteur Myers Diff intégré à l'Overlay WPF. | Clarté maximale des suppressions/insertions. | **Decided** |
| **DEC-023** | UI Word | Taskpane latérale intégrée dans le ruban Office. | Point d'ancrage permanent et rassurant pour l'utilisateur. | **Decided** |
| **DEC-024** | Port Réseau | Port fixe `5284` sur `127.0.0.1`. | Simplifie la configuration et le manifest de l'Add-in. | **Decided** |
| **DEC-025** | Déploiement | Sideloading du manifest XML au MVP. | Déploiement rapide sans attendre l'accord Global Admin M365. | **Decided** |

---

## 46. MATRICE DE TRAÇABILITÉ BOUT-EN-BOUT

| Business Goal | Exigence Fonctionnelle | User Story | Composant Logiciel | Critère d'Acceptation | Test Case UAT | Version Cible |
|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| **BG-01** (Qualité & Rapidité) | **FR-030**, **FR-031** | **US-001** | Agent Correction | **CA-01** | **TEST-UAT-03** | MVP |
| **BG-01** (Adaptabilité) | **FR-040**, **FR-041** | **US-002** | Agent Reformulation | **CA-01** | **TEST-UAT-03** | MVP |
| **BG-02** (Bilinguisme) | **FR-050**, **FR-051** | **US-003** | Agent Traduction | **CA-01** | **TEST-UAT-03** | MVP |
| **BG-01** (Fiabilité Chiffrée) | **FR-060**, **FR-061** | **US-004** | Agent Cohérence Data | **CA-01** | **TEST-UAT-03** | MVP |
| **BG-03** (Sécurité & DLP) | **DLP-001**, **DLP-002** | **US-005** | Moteur DLP Local | **CA-02** | **TEST-UAT-03** | MVP |
| **BG-04** (Ergonomie & Workflow)| **FR-001**, **FR-080** | **US-001** | Overlay WPF & Add-in | **CA-01** | **TEST-UAT-02** | MVP |
| **BG-03** (Résilience Réseau) | **AI-006**, **SEC-005** | **US-005** | LLM Gateway / Failover | **CA-01** | **TEST-UAT-04** | MVP |
| **BG-04** (Déploiement Pilote) | **ADM-006**, **FR-100** | **US-005** | Package MSI / Sideload | **CA-01** | **TEST-UAT-01** | MVP |

---

### Fin du Dossier de Spécifications — ALENIA v1.2 Consolidée