# Indian Express Daily Article Counter & Analyzer

Yeh tool **The Indian Express (`indianexpress.com`)** ke daily official XML sitemaps se articles ko scan aur count karta hai.

---

## 🌐 1. Web App / Browser URL (Naya Feature!)

Aap isko direct browser mein ek **Web Dashboard** ya **Direct URL** ki tarah chala sakte hain:

### Web App Start Karein:
- **`start_web_app.bat`** par double click karein.
- Browser automatically open ho jayega: **`http://localhost:5000/`**

### Direct URL Examples (Dates URL me daal kar data nikalna):

1. **Dashboard me dates auto-load karke direct dekhna:**
   ```text
   http://localhost:5000/?start=2026-11-01&end=2026-11-15
   ```
   *(Is URL ko open karte hi dates auto-fill ho jayengi aur sara data table & cards me screen par dikh jayega)*

2. **Direct Excel File Download URL:**
   ```text
   http://localhost:5000/download?start=2026-11-01&end=2026-11-15
   ```
   *(Browser me enter dabate hi formatted .xlsx report direct download hona shuru ho jayegi)*

3. **Direct JSON API (Integration / Automation ke liye):**
   ```text
   http://localhost:5000/api/count?start=2026-11-01&end=2026-11-15
   ```

---

## 💻 2. Command Line / Batch se Chalana

### Double Click (Terminal Mode):
- **`run_counter.bat`** par double-click karein aur Start/End dates type karein.

### Python Command:
```bash
python count_articles.py --start 2026-11-01 --end 2026-11-15
```

---

## 📊 Kya-kya Data Milega?

1. **Section-wise Count**: Cities, Entertainment, India, Sports, Lifestyle, Opinion, Explained, Business, etc.
2. **Daily Total Count**: Har tareekh ke total articles aur day of the week.
3. **Monthly Count**: Month-wise total articles, active days count aur daily average.
4. **Daily Section Matrix**: Har tareekh ko kis section me kitne articles aaye (Date x Section pivot table).
5. **Excel Report**: 5 sheets me professional styling aur clickable article links ke sath.
