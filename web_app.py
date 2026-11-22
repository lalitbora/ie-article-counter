#!/usr/bin/env python3
"""
The Indian Express — Story Count Web Dashboard & API Server
============================================================
A full-featured web application and local server featuring:
  - Checkbox selectors for Daily, Weekly, Monthly, Section-wise, and Full Metadata
  - Start Date and End Date range pickers with quick shortcuts
  - Real-time tabular metrics, charts, and pivot matrices
  - One-click Excel (.xlsx) report generation and download
  - REST API endpoint (/api/count) and direct download endpoint (/download)

Author: Antigravity Assistant
Environment: Python 3.8+
"""

import sys
import os
import io
import json
import urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import datetime, timedelta
import pandas as pd

# Import core extraction logic from count_articles.py
from count_articles import (
    fetch_all_dates,
    parse_flexible_date,
    export_to_excel,
    PROPERTIES
)

PORT = 5000

HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>The Indian Express — Story Count & Article Analyzer</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {
      --primary: #8b0000;
      --primary-hover: #6e0000;
      --navy: #1F4E78;
      --bg: #F4F6F9;
      --card-bg: #FFFFFF;
      --text: #2D3748;
      --text-muted: #718096;
      --border: #E2E8F0;
      --accent: #3182CE;
      --radius: 10px;
      --shadow: 0 4px 6px -1px rgba(0,0,0,0.07), 0 2px 4px -1px rgba(0,0,0,0.04);
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
      background-color: var(--bg);
      color: var(--text);
      line-height: 1.5;
      padding-bottom: 60px;
    }

    header {
      background: linear-gradient(135deg, #1a202c 0%, #2d3748 100%);
      color: white;
      padding: 22px 20px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.15);
      border-bottom: 4px solid var(--primary);
    }
    .header-inner {
      max-width: 1240px;
      margin: 0 auto;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 12px;
    }
    .header-title h1 {
      font-size: 1.55rem;
      font-weight: 700;
      letter-spacing: -0.5px;
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .badge-ie {
      background: var(--primary);
      color: white;
      font-size: 0.75rem;
      font-weight: 700;
      padding: 3px 8px;
      border-radius: 4px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .header-subtitle {
      font-size: 0.85rem;
      color: #A0AEC0;
      margin-top: 4px;
    }

    .container {
      max-width: 1240px;
      margin: 24px auto;
      padding: 0 20px;
    }

    /* Form Card */
    .card {
      background: var(--card-bg);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      padding: 24px;
      margin-bottom: 24px;
      border: 1px solid var(--border);
    }

    /* Top Module Checkboxes */
    .checkbox-panel {
      background: #F8FAFC;
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 16px 20px;
      margin-bottom: 20px;
    }
    .checkbox-title {
      font-size: 0.85rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: var(--navy);
      margin-bottom: 12px;
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .checkbox-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
      gap: 12px;
    }
    .checkbox-item {
      display: flex;
      align-items: center;
      gap: 8px;
      cursor: pointer;
      font-size: 0.9rem;
      font-weight: 500;
      color: var(--text);
      user-select: none;
    }
    .checkbox-item input[type="checkbox"] {
      width: 18px;
      height: 18px;
      cursor: pointer;
      accent-color: var(--navy);
    }

    .form-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 18px;
      align-items: flex-end;
    }
    .form-group label {
      display: block;
      font-size: 0.85rem;
      font-weight: 600;
      color: var(--text);
      margin-bottom: 6px;
    }
    .form-group input[type="date"] {
      width: 100%;
      padding: 10px 14px;
      font-size: 0.95rem;
      border: 1.5px solid var(--border);
      border-radius: 6px;
      background: #FAFAFA;
      transition: all 0.2s;
    }
    .form-group input[type="date"]:focus {
      outline: none;
      border-color: var(--navy);
      background: #FFFFFF;
      box-shadow: 0 0 0 3px rgba(31, 78, 120, 0.15);
    }

    .btn-group {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
    }
    .btn {
      padding: 11px 22px;
      font-size: 0.95rem;
      font-weight: 600;
      border-radius: 6px;
      cursor: pointer;
      border: none;
      transition: all 0.2s;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      text-decoration: none;
    }
    .btn-primary {
      background: var(--navy);
      color: white;
    }
    .btn-primary:hover {
      background: #153654;
      box-shadow: 0 4px 10px rgba(31, 78, 120, 0.3);
    }
    .btn-success {
      background: #2F855A;
      color: white;
    }
    .btn-success:hover {
      background: #276749;
      box-shadow: 0 4px 10px rgba(47, 133, 90, 0.3);
    }
    .btn:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }

    /* Shortcuts */
    .shortcuts {
      display: flex;
      gap: 8px;
      margin-top: 14px;
      flex-wrap: wrap;
      align-items: center;
    }
    .shortcuts-label {
      font-size: 0.8rem;
      color: var(--text-muted);
      font-weight: 500;
    }
    .shortcut-btn {
      background: #EDF2F7;
      color: #4A5568;
      border: none;
      padding: 4px 10px;
      font-size: 0.8rem;
      border-radius: 4px;
      cursor: pointer;
      font-weight: 500;
      transition: all 0.15s;
    }
    .shortcut-btn:hover {
      background: #CBD5E0;
      color: #2D3748;
    }

    /* KPI Cards */
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 18px;
      margin-bottom: 24px;
    }
    .kpi-card {
      background: var(--card-bg);
      border-radius: var(--radius);
      padding: 20px;
      box-shadow: var(--shadow);
      border: 1px solid var(--border);
      border-left: 5px solid var(--navy);
    }
    .kpi-card:nth-child(2) { border-left-color: #2F855A; }
    .kpi-card:nth-child(3) { border-left-color: #D69E2E; }
    .kpi-card:nth-child(4) { border-left-color: var(--primary); }
    .kpi-label {
      font-size: 0.8rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: var(--text-muted);
      margin-bottom: 6px;
    }
    .kpi-val {
      font-size: 1.9rem;
      font-weight: 700;
      color: var(--text);
      line-height: 1.1;
    }
    .kpi-sub {
      font-size: 0.8rem;
      color: var(--text-muted);
      margin-top: 6px;
    }

    /* Tabs */
    .tabs-nav {
      display: flex;
      gap: 6px;
      border-bottom: 2px solid var(--border);
      margin-bottom: 20px;
      flex-wrap: wrap;
    }
    .tab-btn {
      padding: 10px 18px;
      font-size: 0.9rem;
      font-weight: 600;
      color: var(--text-muted);
      background: none;
      border: none;
      cursor: pointer;
      border-bottom: 3px solid transparent;
      margin-bottom: -2px;
      transition: all 0.2s;
    }
    .tab-btn:hover {
      color: var(--navy);
    }
    .tab-btn.active {
      color: var(--navy);
      border-bottom-color: var(--navy);
    }
    .tab-content {
      display: none;
    }
    .tab-content.active {
      display: block;
    }

    /* Tables */
    .table-container {
      overflow-x: auto;
      border: 1px solid var(--border);
      border-radius: 8px;
      background: white;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.88rem;
      text-align: left;
    }
    th {
      background: #F7FAFC;
      color: #4A5568;
      font-weight: 600;
      padding: 12px 16px;
      border-bottom: 2px solid var(--border);
      white-space: nowrap;
    }
    td {
      padding: 11px 16px;
      border-bottom: 1px solid var(--border);
      color: var(--text);
    }
    tr:last-child td {
      border-bottom: none;
    }
    tr.total-row td {
      background: #EDF2F7;
      font-weight: 700;
      border-top: 2px solid #CBD5E0;
    }
    .text-right { text-align: right; }
    .text-center { text-align: center; }

    .bar-cell {
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .progress-bar {
      flex: 1;
      height: 8px;
      background: #EDF2F7;
      border-radius: 4px;
      overflow: hidden;
      min-width: 60px;
    }
    .progress-fill {
      height: 100%;
      background: var(--navy);
      border-radius: 4px;
    }

    /* Loading Spinner */
    .loading-overlay {
      display: none;
      text-align: center;
      padding: 50px 20px;
    }
    .spinner {
      border: 4px solid #EDF2F7;
      border-top: 4px solid var(--navy);
      border-radius: 50%;
      width: 44px;
      height: 44px;
      animation: spin 0.8s linear infinite;
      margin: 0 auto 16px;
    }
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }

    .url-tip {
      background: #EBF8FF;
      border-left: 4px solid #3182CE;
      padding: 14px 18px;
      border-radius: 6px;
      margin-top: 20px;
      font-size: 0.85rem;
      color: #2B6CB0;
    }
    .url-tip code {
      background: #BEE3F8;
      padding: 2px 6px;
      border-radius: 4px;
      font-family: monospace;
      color: #1A365D;
    }
  </style>
</head>
<body>

  <header>
    <div class="header-inner">
      <div class="header-title">
        <span class="badge-ie">Official Sitemaps</span>
        <h1>Indian Express Group &mdash; Story Count & Article Analyzer</h1>
      </div>
      <div class="header-subtitle">
        Multi-Property Story Counter (The Indian Express, Financial Express, Loksatta, Jansatta, Tamil, Bengali, Gujarati, Malayalam)
      </div>
    </div>
  </header>

  <div class="container">
    <div class="card">
      <form id="counterForm" onsubmit="handleCalculate(event)">
        <!-- Top Options Checkboxes as requested -->
        <div class="checkbox-panel">
          <div class="checkbox-title">
            <span>⚙️ Select Story Count Modules to Include:</span>
            <div style="margin-left: auto; display: flex; gap: 8px;">
              <button type="button" class="shortcut-btn" onclick="toggleAllBoxes(true)">✓ Select All</button>
              <button type="button" class="shortcut-btn" onclick="toggleAllBoxes(false)">✗ Clear All</button>
            </div>
          </div>
          <div class="checkbox-grid">
            <label class="checkbox-item">
              <input type="checkbox" id="chkDaily" checked>
              <span>Daily Story Count (Date-wise)</span>
            </label>
            <label class="checkbox-item">
              <input type="checkbox" id="chkWeekly" checked>
              <span>Weekly Story Count (Week-wise)</span>
            </label>
            <label class="checkbox-item">
              <input type="checkbox" id="chkMonthly" checked>
              <span>Monthly Story Count (Month-wise)</span>
            </label>
            <label class="checkbox-item">
              <input type="checkbox" id="chkSections" checked>
              <span>Section-Wise Breakdown</span>
            </label>
            <label class="checkbox-item">
              <input type="checkbox" id="chkDetails" checked>
              <span>All Metadata (URLs, Headlines, Times)</span>
            </label>
          </div>
        </div>

        <!-- Property Selection Dropdown -->
        <div style="margin-bottom: 20px;">
          <label for="propertySelect" style="display: block; font-size: 0.85rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: var(--navy); margin-bottom: 8px;">
            🌐 Select Publication / Property:
          </label>
          <select id="propertySelect" name="property" style="width: 100%; padding: 11px 14px; border-radius: 6px; border: 1.5px solid var(--navy); font-size: 1rem; font-weight: 600; color: #1F4E78; background: #F8FAFC; cursor: pointer;">
            <option value="indianexpress" selected>📰 The Indian Express (indianexpress.com)</option>
            <option value="financialexpress">📈 Financial Express (financialexpress.com)</option>
            <option value="loksatta">🗞️ Loksatta (loksatta.com)</option>
            <option value="jansatta">🇮🇳 Jansatta (jansatta.com)</option>
            <option value="ietamil">🏛️ IE Tamil (tamil.indianexpress.com)</option>
            <option value="iebengali">✒️ IE Bengali (bengali.indianexpress.com)</option>
            <option value="iegujarati">🦁 IE Gujarati (gujarati.indianexpress.com)</option>
            <option value="iemalayalam">🌴 IE Malayalam (malayalam.indianexpress.com)</option>
          </select>
        </div>

        <div class="form-grid">
          <div class="form-group">
            <label for="startDate">Start Date</label>
            <input type="date" id="startDate" name="startDate" required>
          </div>
          <div class="form-group">
            <label for="endDate">End Date</label>
            <input type="date" id="endDate" name="endDate" required>
          </div>
          <div class="form-group btn-group">
            <button type="submit" id="btnCalc" class="btn btn-primary">
              <span>📊 Calculate & Analyze</span>
            </button>
            <button type="button" id="btnDownload" onclick="handleDownload()" class="btn btn-success" disabled>
              <span>📥 Download Excel (.xlsx)</span>
            </button>
          </div>
        </div>

        <div class="shortcuts">
          <span class="shortcuts-label">Quick Select:</span>
          <button type="button" class="shortcut-btn" onclick="setDates(0, 0)">Today</button>
          <button type="button" class="shortcut-btn" onclick="setDates(1, 1)">Yesterday</button>
          <button type="button" class="shortcut-btn" onclick="setDates(6, 0)">Last 7 Days</button>
          <button type="button" class="shortcut-btn" onclick="setDates(13, 0)">Last 14 Days</button>
          <button type="button" class="shortcut-btn" onclick="setMonthToDate()">This Month</button>
        </div>
      </form>

      <div class="url-tip">
        💡 <strong>Direct URL Link Feature:</strong> Bookmark or share direct links with publication, dates & modules:<br>
        <code>http://localhost:5000/?property=financialexpress&start=2024-12-01&end=2024-12-02</code><br>
        <code>http://localhost:5000/download?property=loksatta&start=2024-12-01&end=2024-12-02</code> &mdash; Direct Excel report download.
      </div>
    </div>

    <!-- Loading Spinner -->
    <div id="loading" class="card loading-overlay">
      <div class="spinner"></div>
      <h3 style="font-size: 1.15rem; font-weight: 600;">Fetching Stories from Official Sitemaps...</h3>
      <p style="color: var(--text-muted); font-size: 0.85rem; margin-top: 6px;">
        Concurrently parsing daily XML feeds and calculating Daily, Weekly, and Monthly counts...
      </p>
    </div>

    <!-- Results Section -->
    <div id="results" style="display: none;">
      <!-- KPI Cards -->
      <div class="kpi-grid">
        <div class="kpi-card">
          <div class="kpi-label">Total Stories</div>
          <div class="kpi-val" id="kpiTotal">0</div>
          <div class="kpi-sub" id="kpiPeriod">Selected Date Range</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-label">Daily Average</div>
          <div class="kpi-val" id="kpiDailyAvg">0</div>
          <div class="kpi-sub">Stories per day</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-label">Total Days</div>
          <div class="kpi-val" id="kpiDays">0</div>
          <div class="kpi-sub">Calendar days analyzed</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-label">Top Desk / Section</div>
          <div class="kpi-val" id="kpiTopSection" style="font-size: 1.45rem;">-</div>
          <div class="kpi-sub" id="kpiTopSectionCount">Highest publishing desk</div>
        </div>
      </div>

      <!-- Detail Card with Tabs -->
      <div class="card">
        <div class="tabs-nav" id="tabsNav">
          <button class="tab-btn active" id="btnTabDaily" onclick="switchTab('tabDaily')">📅 Daily Story Count</button>
          <button class="tab-btn" id="btnTabWeekly" onclick="switchTab('tabWeekly')">🗓️ Weekly Story Count</button>
          <button class="tab-btn" id="btnTabMonthly" onclick="switchTab('tabMonthly')">📆 Monthly Summary</button>
          <button class="tab-btn" id="btnTabSection" onclick="switchTab('tabSection')">📊 Section-Wise Breakdown</button>
          <button class="tab-btn" id="btnTabMatrix" onclick="switchTab('tabMatrix')">🔲 Daily Section Matrix</button>
          <button class="tab-btn" id="btnTabDetails" onclick="switchTab('tabDetails')">📄 All Articles Metadata</button>
        </div>

        <!-- TAB 1: Daily Counts -->
        <div id="tabDaily" class="tab-content active">
          <div class="table-container">
            <table>
              <thead>
                <tr>
                  <th class="text-center" style="width: 140px;">Publish Date</th>
                  <th style="width: 150px;">Day of Week</th>
                  <th class="text-right" style="width: 150px;">Total Stories Published</th>
                  <th>Distribution Chart</th>
                </tr>
              </thead>
              <tbody id="tbodyDaily"></tbody>
            </table>
          </div>
        </div>

        <!-- TAB 2: Weekly Counts -->
        <div id="tabWeekly" class="tab-content">
          <div class="table-container">
            <table>
              <thead>
                <tr>
                  <th>Calendar Week</th>
                  <th class="text-right" style="width: 140px;">Days in Range</th>
                  <th class="text-right" style="width: 140px;">Total Stories</th>
                  <th class="text-right" style="width: 140px;">Daily Average</th>
                  <th>Top Publishing Section</th>
                </tr>
              </thead>
              <tbody id="tbodyWeekly"></tbody>
            </table>
          </div>
        </div>

        <!-- TAB 3: Monthly Summary -->
        <div id="tabMonthly" class="tab-content">
          <div class="table-container">
            <table>
              <thead>
                <tr>
                  <th>Month</th>
                  <th class="text-right">Days Active</th>
                  <th class="text-right">Total Stories</th>
                  <th class="text-right">Daily Average</th>
                </tr>
              </thead>
              <tbody id="tbodyMonthly"></tbody>
            </table>
          </div>
        </div>

        <!-- TAB 4: Section Summary -->
        <div id="tabSection" class="tab-content">
          <div class="table-container">
            <table>
              <thead>
                <tr>
                  <th class="text-center" style="width: 60px;">Rank</th>
                  <th>Section / Desk Name</th>
                  <th class="text-right" style="width: 140px;">Story Count</th>
                  <th style="width: 250px;">Share (%)</th>
                </tr>
              </thead>
              <tbody id="tbodySection"></tbody>
            </table>
          </div>
        </div>

        <!-- TAB 5: Daily Section Matrix -->
        <div id="tabMatrix" class="tab-content">
          <div class="table-container" style="max-height: 520px; overflow-y: auto;">
            <table id="tableMatrix">
              <thead id="theadMatrix"></thead>
              <tbody id="tbodyMatrix"></tbody>
            </table>
          </div>
        </div>

        <!-- TAB 6: All Articles Metadata -->
        <div id="tabDetails" class="tab-content">
          <div style="margin-bottom: 12px; font-size: 0.85rem; color: var(--text-muted);">
            Showing all published stories with Headlines, First Publish Time, Last Modified Time, Section and URLs:
          </div>
          <div class="table-container" style="max-height: 550px; overflow-y: auto;">
            <table>
              <thead>
                <tr>
                  <th class="text-center" style="width: 50px;">#</th>
                  <th class="text-center" style="width: 100px;">Date</th>
                  <th style="width: 100px;">Day</th>
                  <th class="text-center" style="width: 90px;">Time (IST)</th>
                  <th style="width: 130px;">Section</th>
                  <th style="width: 130px;">Sub-Section</th>
                  <th>Headline / Title</th>
                  <th style="width: 90px;">Link</th>
                </tr>
              </thead>
              <tbody id="tbodyDetails"></tbody>
            </table>
          </div>
        </div>

      </div>
    </div>
  </div>

  <script>
    function toggleAllBoxes(state) {
      document.getElementById('chkDaily').checked = state;
      document.getElementById('chkWeekly').checked = state;
      document.getElementById('chkMonthly').checked = state;
      document.getElementById('chkSections').checked = state;
      document.getElementById('chkDetails').checked = state;
    }

    function formatDate(d) {
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }

    function setDates(startOffsetDays, endOffsetDays) {
      const now = new Date();
      const end = new Date(now);
      end.setDate(end.getDate() - endOffsetDays);
      const start = new Date(now);
      start.setDate(start.getDate() - startOffsetDays);

      document.getElementById('startDate').value = formatDate(start);
      document.getElementById('endDate').value = formatDate(end);
    }

    function setMonthToDate() {
      const now = new Date();
      const start = new Date(now.getFullYear(), now.getMonth(), 1);
      document.getElementById('startDate').value = formatDate(start);
      document.getElementById('endDate').value = formatDate(now);
    }

    function switchTab(tabId) {
      document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));

      event.target.classList.add('active');
      document.getElementById(tabId).classList.add('active');
    }

    async function handleCalculate(e) {
      if (e) e.preventDefault();

      const prop = document.getElementById('propertySelect').value;
      const start = document.getElementById('startDate').value;
      const end = document.getElementById('endDate').value;

      if (!start || !end) {
        alert('Please select both Start Date and End Date.');
        return;
      }

      const incDaily = document.getElementById('chkDaily').checked ? '1' : '0';
      const incWeekly = document.getElementById('chkWeekly').checked ? '1' : '0';
      const incMonthly = document.getElementById('chkMonthly').checked ? '1' : '0';
      const incSections = document.getElementById('chkSections').checked ? '1' : '0';
      const incDetails = document.getElementById('chkDetails').checked ? '1' : '0';

      // Update URL query parameters
      const url = new URL(window.location);
      url.searchParams.set('property', prop);
      url.searchParams.set('start', start);
      url.searchParams.set('end', end);
      url.searchParams.set('daily', incDaily);
      url.searchParams.set('weekly', incWeekly);
      url.searchParams.set('monthly', incMonthly);
      url.searchParams.set('sections', incSections);
      url.searchParams.set('details', incDetails);
      window.history.pushState({}, '', url);

      document.getElementById('loading').style.display = 'block';
      document.getElementById('results').style.display = 'none';
      document.getElementById('btnCalc').disabled = true;
      document.getElementById('btnDownload').disabled = true;

      try {
        const queryStr = `/api/count?property=${encodeURIComponent(prop)}&start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&daily=${incDaily}&weekly=${incWeekly}&monthly=${incMonthly}&sections=${incSections}&details=${incDetails}`;
        const resp = await fetch(queryStr);
        const data = await resp.json();

        if (data.status !== 'success') {
          alert('Error: ' + (data.message || 'Failed to fetch article counts.'));
          return;
        }

        renderDashboard(data);
        document.getElementById('btnDownload').disabled = false;
      } catch (err) {
        alert('Connection error: ' + err.message);
      } finally {
        document.getElementById('loading').style.display = 'none';
        document.getElementById('btnCalc').disabled = false;
      }
    }

    function renderDashboard(data) {
      document.getElementById('results').style.display = 'block';

      // KPI Top Cards
      document.getElementById('kpiTotal').textContent = Number(data.total_articles).toLocaleString();
      document.getElementById('kpiDailyAvg').textContent = data.daily_average.toFixed(1);
      document.getElementById('kpiDays').textContent = data.total_days;
      document.getElementById('kpiPeriod').textContent = `${data.start_date} to ${data.end_date}`;

      if (data.section_summary && data.section_summary.length > 0) {
        const top = data.section_summary[0];
        document.getElementById('kpiTopSection').textContent = top.section;
        document.getElementById('kpiTopSectionCount').textContent = `${Number(top.count).toLocaleString()} stories (${top.share_pct}%)`;
      }

      // 1. Daily Table
      const tbodyDaily = document.getElementById('tbodyDaily');
      let dHtml = '';
      if (data.daily_counts && data.daily_counts.length > 0) {
        const maxDaily = Math.max(...data.daily_counts.map(d => d.total_articles), 1);
        data.daily_counts.forEach(d => {
          const pct = Math.round((d.total_articles / maxDaily) * 100);
          dHtml += `
            <tr>
              <td class="text-center"><code>${d.date}</code></td>
              <td>${d.day}</td>
              <td class="text-right"><strong>${Number(d.total_articles).toLocaleString()}</strong></td>
              <td>
                <div class="bar-cell">
                  <div class="progress-bar">
                    <div class="progress-fill" style="width: ${pct}%;"></div>
                  </div>
                  <span style="font-size: 0.8rem; color: var(--text-muted); width: 40px;">${d.total_articles}</span>
                </div>
              </td>
            </tr>
          `;
        });
        dHtml += `
          <tr class="total-row">
            <td class="text-center">GRAND TOTAL</td>
            <td>${data.total_days} Days</td>
            <td class="text-right">${Number(data.total_articles).toLocaleString()}</td>
            <td>-</td>
          </tr>
        `;
      } else {
        dHtml = '<tr><td colspan="4" class="text-center">Daily module was not selected.</td></tr>';
      }
      tbodyDaily.innerHTML = dHtml;

      // 2. Weekly Table
      const tbodyWeekly = document.getElementById('tbodyWeekly');
      let wHtml = '';
      if (data.weekly_summary && data.weekly_summary.length > 0) {
        data.weekly_summary.forEach(w => {
          wHtml += `
            <tr>
              <td><strong>${w.week_label}</strong></td>
              <td class="text-right">${w.days_in_range}</td>
              <td class="text-right"><strong>${Number(w.total_stories).toLocaleString()}</strong></td>
              <td class="text-right">${w.daily_average.toFixed(1)}</td>
              <td>${w.top_section}</td>
            </tr>
          `;
        });
        wHtml += `
          <tr class="total-row">
            <td>GRAND TOTAL</td>
            <td class="text-right">${data.total_days}</td>
            <td class="text-right">${Number(data.total_articles).toLocaleString()}</td>
            <td class="text-right">${data.daily_average.toFixed(1)}</td>
            <td>All Sections</td>
          </tr>
        `;
      } else {
        wHtml = '<tr><td colspan="5" class="text-center">Weekly module was not selected.</td></tr>';
      }
      tbodyWeekly.innerHTML = wHtml;

      // 3. Monthly Table
      const tbodyMonthly = document.getElementById('tbodyMonthly');
      let mHtml = '';
      if (data.monthly_summary && data.monthly_summary.length > 0) {
        data.monthly_summary.forEach(m => {
          mHtml += `
            <tr>
              <td><strong>${m.month}</strong></td>
              <td class="text-right">${m.days_active}</td>
              <td class="text-right"><strong>${Number(m.total_articles).toLocaleString()}</strong></td>
              <td class="text-right">${m.daily_average.toFixed(1)}</td>
            </tr>
          `;
        });
        mHtml += `
          <tr class="total-row">
            <td>TOTAL</td>
            <td class="text-right">${data.total_days}</td>
            <td class="text-right">${Number(data.total_articles).toLocaleString()}</td>
            <td class="text-right">${data.daily_average.toFixed(1)}</td>
          </tr>
        `;
      } else {
        mHtml = '<tr><td colspan="4" class="text-center">Monthly module was not selected.</td></tr>';
      }
      tbodyMonthly.innerHTML = mHtml;

      // 4. Section Table
      const tbodySection = document.getElementById('tbodySection');
      let sHtml = '';
      if (data.section_summary && data.section_summary.length > 0) {
        data.section_summary.forEach(s => {
          sHtml += `
            <tr>
              <td class="text-center">${s.rank}</td>
              <td><strong>${s.section}</strong></td>
              <td class="text-right"><strong>${Number(s.count).toLocaleString()}</strong></td>
              <td>
                <div class="bar-cell">
                  <div class="progress-bar">
                    <div class="progress-fill" style="width: ${Math.min(s.share_pct * 3, 100)}%;"></div>
                  </div>
                  <span style="font-size: 0.85rem; font-weight: 600; width: 50px;">${s.share_pct}%</span>
                </div>
              </td>
            </tr>
          `;
        });
        sHtml += `
          <tr class="total-row">
            <td class="text-center">-</td>
            <td>GRAND TOTAL</td>
            <td class="text-right">${Number(data.total_articles).toLocaleString()}</td>
            <td>100.0%</td>
          </tr>
        `;
      } else {
        sHtml = '<tr><td colspan="4" class="text-center">Section breakdown was not selected.</td></tr>';
      }
      tbodySection.innerHTML = sHtml;

      // 5. Daily Section Matrix
      renderMatrix(data.matrix);

      // 6. Full Metadata (All Articles)
      const tbodyDetails = document.getElementById('tbodyDetails');
      let dtlHtml = '';
      if (data.all_articles && data.all_articles.length > 0) {
        data.all_articles.forEach((art, idx) => {
          dtlHtml += `
            <tr>
              <td class="text-center">${idx + 1}</td>
              <td class="text-center"><code>${art.date}</code></td>
              <td>${art.day_of_week}</td>
              <td class="text-center">${art.first_publish_time}</td>
              <td><strong>${art.section}</strong></td>
              <td>${art.sub_section}</td>
              <td>${art.headline}</td>
              <td><a href="${art.url}" target="_blank" style="color: #0563C1; font-weight: 600; text-decoration: underline;">View URL</a></td>
            </tr>
          `;
        });
      } else {
        dtlHtml = '<tr><td colspan="8" class="text-center">Article metadata was not requested or no articles found.</td></tr>';
      }
      tbodyDetails.innerHTML = dtlHtml;
    }

    function renderMatrix(matrix) {
      const thead = document.getElementById('theadMatrix');
      const tbody = document.getElementById('tbodyMatrix');

      if (!matrix || !matrix.sections || matrix.sections.length === 0) {
        thead.innerHTML = '<tr><th>Date</th><th>Status</th></tr>';
        tbody.innerHTML = '<tr><td colspan="2" class="text-center">Matrix data not available.</td></tr>';
        return;
      }

      let headHtml = '<tr><th style="position: sticky; left: 0; background: #F7FAFC; z-index: 2;">Date</th>';
      matrix.sections.forEach(sec => {
        headHtml += `<th class="text-right">${sec}</th>`;
      });
      headHtml += '<th class="text-right" style="background: #EDF2F7;">Total</th></tr>';
      thead.innerHTML = headHtml;

      let bodyHtml = '';
      matrix.rows.forEach(r => {
        const isTotal = (r.date === 'Grand Total');
        const rowClass = isTotal ? 'class="total-row"' : '';
        const stickyStyle = isTotal ? 'position: sticky; left: 0; background: #EDF2F7; z-index: 1;' : 'position: sticky; left: 0; background: white; z-index: 1;';
        
        bodyHtml += `<tr ${rowClass}>`;
        bodyHtml += `<td style="${stickyStyle}"><strong>${r.date}</strong></td>`;
        matrix.sections.forEach(sec => {
          const val = r.counts[sec] || 0;
          bodyHtml += `<td class="text-right">${val ? Number(val).toLocaleString() : '-'}</td>`;
        });
        bodyHtml += `<td class="text-right" style="font-weight: 700;">${Number(r.total).toLocaleString()}</td>`;
        bodyHtml += '</tr>';
      });
      tbody.innerHTML = bodyHtml;
    }

    function handleDownload() {
      const prop = document.getElementById('propertySelect').value;
      const start = document.getElementById('startDate').value;
      const end = document.getElementById('endDate').value;
      if (!start || !end) return;

      const incDaily = document.getElementById('chkDaily').checked ? '1' : '0';
      const incWeekly = document.getElementById('chkWeekly').checked ? '1' : '0';
      const incMonthly = document.getElementById('chkMonthly').checked ? '1' : '0';
      const incSections = document.getElementById('chkSections').checked ? '1' : '0';
      const incDetails = document.getElementById('chkDetails').checked ? '1' : '0';

      window.location.href = `/download?property=${encodeURIComponent(prop)}&start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&daily=${incDaily}&weekly=${incWeekly}&monthly=${incMonthly}&sections=${incSections}&details=${incDetails}`;
    }

    // Auto-fill from query params on page load
    window.addEventListener('DOMContentLoaded', () => {
      const params = new URLSearchParams(window.location.search);
      const prop = params.get('property');
      const start = params.get('start');
      const end = params.get('end');

      if (prop && document.getElementById('propertySelect')) {
        document.getElementById('propertySelect').value = prop;
      }

      if (params.has('daily')) document.getElementById('chkDaily').checked = (params.get('daily') === '1');
      if (params.has('weekly')) document.getElementById('chkWeekly').checked = (params.get('weekly') === '1');
      if (params.has('monthly')) document.getElementById('chkMonthly').checked = (params.get('monthly') === '1');
      if (params.has('sections')) document.getElementById('chkSections').checked = (params.get('sections') === '1');
      if (params.has('details')) document.getElementById('chkDetails').checked = (params.get('details') === '1');

      if (start && end) {
        document.getElementById('startDate').value = start;
        document.getElementById('endDate').value = end;
        handleCalculate();
      } else {
        // Default to last 7 days
        setDates(6, 0);
      }
    });
  </script>
</body>
</html>
"""


class ArticleCounterHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        # Clean console logging
        sys.stderr.write(f"[{datetime.now().strftime('%H:%M:%S')}] {args[0]} {args[1]}\n")

    def do_GET(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path
        params = urllib.parse.parse_qs(parsed_url.query)

        # 1. Main Web Page: GET / or GET /?start=...&end=...
        if path == "/" or path == "/index.html":
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(HTML_PAGE.encode("utf-8"))
            return

        # 2. JSON API: GET /api/count?property=...&start=YYYY-MM-DD&end=YYYY-MM-DD
        elif path == "/api/count":
            prop_key = params.get("property", ["indianexpress"])[0]
            if prop_key not in PROPERTIES:
                prop_key = "indianexpress"
            prop_info = PROPERTIES[prop_key]

            start_str = params.get("start", [None])[0]
            end_str = params.get("end", [None])[0]
            inc_daily = params.get("daily", ["1"])[0] == "1"
            inc_weekly = params.get("weekly", ["1"])[0] == "1"
            inc_monthly = params.get("monthly", ["1"])[0] == "1"
            inc_sections = params.get("sections", ["1"])[0] == "1"
            inc_details = params.get("details", ["1"])[0] == "1"

            if not start_str or not end_str:
                self.send_json_response(400, {
                    "status": "error",
                    "message": "Both 'start' and 'end' query parameters are required (e.g. /api/count?property=financialexpress&start=2024-12-01&end=2024-12-02)"
                })
                return

            try:
                start_date = parse_flexible_date(start_str)
                end_date = parse_flexible_date(end_str)
                if start_date > end_date:
                    start_date, end_date = end_date, start_date

                all_articles, _ = fetch_all_dates(start_date, end_date, property_key=prop_key, max_workers=6)

                if not all_articles:
                    self.send_json_response(200, {
                        "status": "success",
                        "property_key": prop_key,
                        "property_name": prop_info["name"],
                        "start_date": str(start_date),
                        "end_date": str(end_date),
                        "total_days": (end_date - start_date).days + 1,
                        "total_articles": 0,
                        "daily_average": 0.0,
                        "monthly_summary": [],
                        "weekly_summary": [],
                        "daily_counts": [],
                        "section_summary": [],
                        "matrix": {"sections": [], "rows": []},
                        "all_articles": []
                    })
                    return

                df = pd.DataFrame(all_articles)
                total_articles = len(df)
                total_days = (end_date - start_date).days + 1
                daily_avg = total_articles / total_days if total_days > 0 else 0

                # 1. Monthly Summary
                monthly_summary = []
                if inc_monthly:
                    for (m_code, m_name), m_df in df.groupby(['month', 'month_name']):
                        m_count = len(m_df)
                        m_days = m_df['date'].nunique()
                        monthly_summary.append({
                            "month": m_name,
                            "days_active": m_days,
                            "total_articles": m_count,
                            "daily_average": round(m_count / m_days, 1) if m_days > 0 else 0
                        })

                # 2. Weekly Summary
                weekly_summary = []
                if inc_weekly:
                    for (w_key, w_label), w_df in df.groupby(['week_key', 'week_label']):
                        w_count = len(w_df)
                        w_days = w_df['date'].nunique()
                        top_sec = w_df['section'].value_counts().index[0] if not w_df.empty else "-"
                        weekly_summary.append({
                            "week_key": w_key,
                            "week_label": w_label,
                            "days_in_range": w_days,
                            "total_stories": w_count,
                            "daily_average": round(w_count / w_days, 1) if w_days > 0 else 0,
                            "top_section": top_sec
                        })

                # 3. Daily Counts
                daily_counts = []
                if inc_daily:
                    daily_groups = df.groupby(['date', 'day_of_week']).size().reset_index(name='count').sort_values('date')
                    for _, r in daily_groups.iterrows():
                        daily_counts.append({
                            "date": r['date'],
                            "day": r['day_of_week'],
                            "total_articles": int(r['count'])
                        })

                # 4. Section Summary
                section_summary = []
                if inc_sections:
                    sec_counts = df['section'].value_counts()
                    for rank, (sec, count) in enumerate(sec_counts.items(), 1):
                        pct = round((count / total_articles) * 100, 1) if total_articles > 0 else 0
                        section_summary.append({
                            "rank": rank,
                            "section": sec,
                            "count": int(count),
                            "share_pct": pct
                        })

                # 5. Matrix Data
                matrix_sections = []
                matrix_rows = []
                if inc_sections:
                    pivot = pd.crosstab(index=df['date'], columns=df['section'], margins=True, margins_name="Grand Total")
                    matrix_sections = [c for c in pivot.columns if c != "Grand Total"]
                    for dt_val, row_data in pivot.iterrows():
                        counts_dict = {sec: int(row_data.get(sec, 0)) for sec in matrix_sections}
                        matrix_rows.append({
                            "date": str(dt_val),
                            "counts": counts_dict,
                            "total": int(row_data["Grand Total"])
                        })

                # 6. Detailed Articles Metadata
                all_articles_out = []
                if inc_details:
                    all_articles_out = df[['date', 'day_of_week', 'first_publish_time', 'last_modified_time', 'section', 'sub_section', 'headline', 'url']].to_dict('records')

                response_payload = {
                    "status": "success",
                    "property_key": prop_key,
                    "property_name": prop_info["name"],
                    "start_date": str(start_date),
                    "end_date": str(end_date),
                    "total_days": total_days,
                    "total_articles": total_articles,
                    "daily_average": round(daily_avg, 1),
                    "monthly_summary": monthly_summary,
                    "weekly_summary": weekly_summary,
                    "daily_counts": daily_counts,
                    "section_summary": section_summary,
                    "matrix": {
                        "sections": matrix_sections,
                        "rows": matrix_rows
                    },
                    "all_articles": all_articles_out
                }
                self.send_json_response(200, response_payload)

            except Exception as e:
                self.send_json_response(500, {
                    "status": "error",
                    "message": str(e)
                })
            return

        # 3. Direct Excel Download: GET /download?property=...&start=...&end=...
        elif path == "/download":
            prop_key = params.get("property", ["indianexpress"])[0]
            if prop_key not in PROPERTIES:
                prop_key = "indianexpress"
            prop_info = PROPERTIES[prop_key]

            start_str = params.get("start", [None])[0]
            end_str = params.get("end", [None])[0]
            inc_daily = params.get("daily", ["1"])[0] == "1"
            inc_weekly = params.get("weekly", ["1"])[0] == "1"
            inc_monthly = params.get("monthly", ["1"])[0] == "1"
            inc_sections = params.get("sections", ["1"])[0] == "1"
            inc_details = params.get("details", ["1"])[0] == "1"

            if not start_str or not end_str:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b"Error: Both start and end parameters are required.")
                return

            try:
                start_date = parse_flexible_date(start_str)
                end_date = parse_flexible_date(end_str)
                if start_date > end_date:
                    start_date, end_date = end_date, start_date

                all_articles, _ = fetch_all_dates(start_date, end_date, property_key=prop_key, max_workers=6)
                if not all_articles:
                    self.send_response(404)
                    self.end_headers()
                    self.wfile.write(f"No articles found for {prop_info['name']} in this date range.".encode("utf-8"))
                    return

                df = pd.DataFrame(all_articles)

                # Generate styled Excel in memory
                buf = io.BytesIO()
                export_to_excel(
                    df, start_date, end_date, buf,
                    property_name=prop_info['name'],
                    include_daily=inc_daily,
                    include_weekly=inc_weekly,
                    include_monthly=inc_monthly,
                    include_sections=inc_sections,
                    include_details=inc_details
                )
                excel_bytes = buf.getvalue()

                clean_name = prop_info['name'].replace(' ', '_')
                filename = f"{clean_name}_Story_Count_{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}.xlsx"

                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Content-Length", str(len(excel_bytes)))
                self.end_headers()
                self.wfile.write(excel_bytes)

            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(f"Server Error: {e}".encode("utf-8"))
            return

        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b"Not Found")

    def send_json_response(self, status_code, data):
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))


def main():
    server_address = ("127.0.0.1", PORT)
    httpd = HTTPServer(server_address, ArticleCounterHandler)
    print("=" * 68)
    print("  THE INDIAN EXPRESS — STORY COUNT SOFTWARE & SERVER RUNNING")
    print("=" * 68)
    print(f"\n  Web Dashboard URL : http://localhost:{PORT}/")
    print(f"  Example Query URL : http://localhost:{PORT}/?start=2026-11-01&end=2026-11-30")
    print(f"  Direct Download   : http://localhost:{PORT}/download?start=2026-11-01&end=2026-11-30")
    print(f"  JSON API URL      : http://localhost:{PORT}/api/count?start=2026-11-01&end=2026-11-30")
    print("\nPress Ctrl+C in this window to stop the server.\n")

    # Automatically open local browser
    try:
        import webbrowser
        webbrowser.open(f"http://localhost:{PORT}/")
    except Exception:
        pass

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server...")
        httpd.server_close()


if __name__ == "__main__":
    main()
