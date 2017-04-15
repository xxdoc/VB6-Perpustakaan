# VB6-Perpustakaan
VB6 &amp; MySQL

Ruang Lingkup :
1. Mencatat peminjaman dan pengembalian buku
2. Mencatat data buku,anggota dan staf perpustakaan
3. Laporan peminjaman, pengembalian


Fitur Utama :
1. Ada nya menu Setting
   Menu Setting memudahkan seorang admin melakukan perubahan terhadap aplikasi tanpa harus
   membuka database.

2. Kunci Akun atau User
   Jika seorang user salah melakukan login sebanyak yang ditentukan, maka akun atau user
   tersebut akan dikunci demi keamanan. Hanya admin yang bisa membuka user tersebut.

3. Buku Terfavorit

Aturan Perpustakaan:
1. Maksimal buku yang dipinjam secara default yaitu 5 buku per-anggota
   Jika anggota ingin meminjam buku lain, maka harus mengembalikan buku yang dipinjam
   terlebih dahulu sebelum jatuh tempo

2. Denda kerusakan buku secara default 30%

3. Satu anggota hanya bisa meminjam 1 judul buku,tidak boleh lebih

4. Jika anggota mempunyai "hutang" pengembalian buku atau belum mengembalikan buku padahal sudah
   melebihi jatuh tempo, anggota tidak boleh meminjam buku lain meskipun jumlah buku yang dipinjam
   belum maksimal.
