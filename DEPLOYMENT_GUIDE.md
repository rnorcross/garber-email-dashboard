# Garber Email Marketing & Advertising Dashboard — Deployment Guide

## Overview
This dashboard pulls data from your published Google Sheet and displays email marketing and advertising results across 7 tabs:

1. **Group Sales Email** — Overall dealer group sales email marketing KPIs & trends
2. **Group Service Email** — Overall dealer group service email marketing KPIs & trends  
3. **Group Advertising** — Overall Google Ads + Facebook Ads performance
4. **Dealership Sales** — Sales email results ranked by dealership
5. **Dealership Service + Google Ads** — Service email + Google Ads by dealership
6. **Dealership Facebook Ads** — Facebook Ad results by dealership
7. **Customer Data** — Customer-specific data from dealership tabs in the spreadsheet

## Deployment Steps (Same as Phone Call Dashboard)

### 1. Create a New GitHub Repository
- Go to github.com → New Repository
- Name it something like `garber-email-dashboard`
- Make it Public
- Click "Create repository"

### 2. Upload the Project Files
Upload these files to your repo root:
- `package.json`
- `index.html`
- `vite.config.js`

Create a `src/` folder and upload:
- `src/main.jsx`
- `src/App.jsx`

### 3. Deploy on Vercel
- Go to vercel.com and sign in with GitHub
- Click "Add New" → "Project"
- Import your new repo
- Framework Preset: **Vite**
- Click **Deploy**

### 4. Verify
- Once deployed, visit your Vercel URL
- The dashboard should load and pull data from your Google Sheet automatically

## Updating the Dashboard
Same workflow as the phone call dashboard:
1. Start a new Claude conversation
2. Upload the current `App.jsx`
3. Describe what changes you need
4. Copy the updated `App.jsx` back to GitHub using the pencil edit button
5. Vercel auto-redeploys

## Google Sheet Requirements
- Your Google Sheet must be **published to the web** as CSV
- The first tab ("Overall Results") should contain all the monthly data columns
- Additional tabs named by dealership will be auto-discovered for the Customer Data tab

## Features
- **Month selector** — Pick any month to view that period's results
- **Month-over-month arrows** — Shows increase/decrease vs. prior month
- **Color-coded badges** — ROI% and Cost/Lead use green/yellow/red indicators
- **Sortable columns** — Click any column header to sort
- **Export options** — JPG, PDF, or ZIP of all tabs
- **Refresh Data** — Pull latest data from the spreadsheet at any time
- **Responsive design** — Works on desktop and tablet
