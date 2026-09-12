# Hafalan Ikhwan

Tracker progres hafalan Al-Qur'an untuk musyrif, koordinator, dan peserta Kelas Ikhwan, Rumah Quran Tj. Priok.

## Aplikasi yang berjalan (live)

https://claude.ai/code/artifact/f63a533e-0a2c-439a-8524-91a1b4810140

Link di atas adalah aplikasi yang sebenarnya dipakai — data peserta dan riwayat hafalan tersimpan di sana secara real-time.

## Tentang repo ini

`index.html` di repo ini adalah **cadangan kode**, bukan situs yang bisa dijalankan sendiri. Fitur penyimpanan data (`window.claude.use("db")`) hanya berfungsi saat dibuka lewat platform Claude Artifact di atas — kalau file ini di-hosting di tempat lain (GitHub Pages, Vercel, dll.), form input dan dashboard tidak akan menyimpan apa pun karena backend database-nya tidak ikut terbawa.

Setiap kali aplikasi live diperbarui, kode di repo ini ikut di-push supaya riwayatnya tetap tersimpan di sini.
