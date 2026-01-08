# Pop Up CSV Generator (GitHub Pages)

Web app statis: upload Excel Master Pop Up -> download ZIP berisi 6 CSV (HOME, HOME-BIZ, FDT, FAT, HOOK, POLE).
Semua proses terjadi di browser (tanpa upload ke server).

## Cara deploy (GitHub Pages)
1. Buat repo baru di GitHub (Public lebih mudah untuk Pages).
2. Upload file-file dalam folder ini: `index.html`, `app.js`, `styles.css`.
3. Buka repo -> Settings -> Pages.
4. Source: Deploy from a branch.
5. Branch: `main` / folder: `/ (root)`.
6. Save. Tunggu link Pages aktif.

## Cara pakai
- Buka link GitHub Pages
- Upload file `.xlsx` (Master Pop Up)
- (Opsional) isi nama area folder ZIP
- Klik Generate ZIP

## Aturan data
- HOME vs HOME-BIZ: dari kolom `HOME/HOME-BIZ`
- FAT: `FAT ID/NETWORK ID` yang mengandung huruf `S` (bisa multi-ID dipisah `&`/koma)
- FDT: `FAT ID/NETWORK ID` yang tidak mengandung `S` (contoh: `FBS05300`)
- POLE: unik berdasarkan `Pole ID (New)`
- HOOK: unik berdasarkan `Clamp_Hook_ID` (abaikan `-`)

Catatan: bila template/kolom berubah, update daftar header di `app.js` pada objek `TPL`.
