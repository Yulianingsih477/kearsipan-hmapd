/* ============================================================
   KONFIGURASI SUPABASE (DATABASE CLOUD ANTAR PERANGKAT)
   Isi URL dan Anon Key dari Project Settings -> API di Supabase.
   Jika masih kosong, sistem otomatis berjalan dalam Mode Lokal.
   ============================================================ */
const SUPABASE_URL = 'https://qtynikkiuhaddnpgfnqk.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_8XcC4JxbrIngXXWDTMOW2A_7pYEe4co';

let supabaseClient = null;

const DEFAULT_KONTAK_ADMIN = {
    nama: 'Admin Kearsipan PT HMAPD',
    wa: ''
};

const DEFAULT_KONTAK_MANAJER = [
    { id: 1, nama: 'Manajer Operasional', wa: '' },
    { id: 2, nama: 'Manajer Keuangan', wa: '' },
    { id: 3, nama: 'Manajer SDM', wa: '' },
    { id: 4, nama: 'Manajer Administrasi', wa: '' },
    { id: 5, nama: 'Manajer Pemasaran', wa: '' },
    { id: 6, nama: 'Manajer Produksi', wa: '' },
    { id: 7, nama: 'Manajer Umum', wa: '' },
    { id: 8, nama: 'Manajer IT', wa: '' }
];

/* ============================================================
   KEARSIPAN DIGITAL PT HMAPD
   JavaScript perbaikan: satu sumber event, tabel, arsip, disposisi.
   Data awal diambil dari Buku Agenda.xlsx (BASM, BASK, Klasifikasi).
   ============================================================ */

const DATA_EXCEL = {
    "suratMasuk": [
        {
            "id": 1,
            "noAgenda": 1,
            "tglTerima": "2026-09-04",
            "noSurat": "004/121.1/PZ/IX/2026",
            "tglSurat": "2026-09-04",
            "pengirim": "PT PUTRI ZALSABILA",
            "perihal": "Permintaan Daftar Harga",
            "isi": "Pengajuan daftar harga resmi, meliputi spesifikasi barang, dan informasi garansi resmi.",
            "kodeArsip": "321",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 2,
            "noAgenda": 2,
            "tglTerima": "2026-09-09",
            "noSurat": "",
            "tglSurat": "2026-09-08",
            "pengirim": "Siti Aisyah",
            "perihal": "Permohonan Cuti Menikah",
            "isi": "Mengajukan permohonan cuti menikah selama 7 (tujuh) hari, terhitung mulai tanggal 15 September 2026 sampai dengan 21 September 2026.",
            "kodeArsip": "113",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 3,
            "noAgenda": 3,
            "tglTerima": "2026-09-12",
            "noSurat": "",
            "tglSurat": "2026-09-12",
            "pengirim": "Indriani",
            "perihal": "Lamaran Pekerjaan - Manajer Pemasaran",
            "isi": "mengajukan diri untuk melamar dan mengisi posisi sebagai Manajer Pemasaran dengan pengalaman kerja selama empat tahun sebagai Supervisor Pemasaran.",
            "kodeArsip": "121",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 4,
            "noAgenda": 4,
            "tglTerima": "2026-09-12",
            "noSurat": "",
            "tglSurat": "2026-09-12",
            "pengirim": "Wahyu Pratama",
            "perihal": "Lamaran Pekerjaan - Manajer Pemasaran",
            "isi": "mengajukan diri untuk melamar dan mengisi posisi sebagai Manajer Pemasaran dengan pengalaman kerja selama lima tahun sebagai Marketing Manager di perusahaan retail",
            "kodeArsip": "121",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 5,
            "noAgenda": 5,
            "tglTerima": "2026-10-06",
            "noSurat": "001/321/MS/X/2026",
            "tglSurat": "2026-10-06",
            "pengirim": "CV MITRA SEJAHTERA",
            "perihal": "Permintaan Penawaran Harga",
            "isi": "Memohon kesediaan untuk mengirimkan daftar harga (penawaran) resmi\nuntuk jenis barang berikut: meja kerja staff, meja kerja manager, kursi kantor staff, kursi kantor\nmanager, dan lemari arsip, lengkap dengan spesifikasi, syarat pembayaran, dan ketentuan\npengiriman.",
            "kodeArsip": "321",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 6,
            "noAgenda": 6,
            "tglTerima": "2026-10-15",
            "noSurat": "002/341/MS/X/2026",
            "tglSurat": "2026-10-14",
            "pengirim": "CV MITRA SEJAHTERA",
            "perihal": "Pesanan Barang",
            "isi": "Pesanan barang berupa Meja Kerja Staff,\nKursi Kantor Staff 10 Unit, \nLemari Arsip/Filing Cabinet 4 Pintu 5 Unit.",
            "kodeArsip": "341",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 7,
            "noAgenda": 7,
            "tglTerima": "2026-10-25",
            "noSurat": "003/381/MS/X/2026",
            "tglSurat": "2026-10-25",
            "pengirim": "CV MITRA SEJAHTERA",
            "perihal": "Klaim atas Barang Pesanan",
            "isi": "Permohonan klaim atas pesanan: dari 10 unit Meja Kerja Staff Minimalis yang dipesan, 2 unit dalam kondisi permukaan tergores dan 1 engsel laci rusak; jumlah Lemari Arsip yang diterima hanya 4 unit dari 5 unit yang dipesan.",
            "kodeArsip": "381",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 8,
            "noAgenda": 8,
            "tglTerima": "2026-10-27",
            "noSurat": "",
            "tglSurat": "2026-10-27",
            "pengirim": "Rahmat Hidayat",
            "perihal": "Permohonan Cuti Sakit",
            "isi": "Mengajukan cuti sakit selama 3 (tiga) hari kerja, 27 s.d. 29 Oktober 2026, disertai surat keterangan dokter.",
            "kodeArsip": "112",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 9,
            "noAgenda": 9,
            "tglTerima": "2026-10-29",
            "noSurat": "004/371/MS/X/2026",
            "tglSurat": "2026-10-28",
            "pengirim": "CV MITRA SEJAHTERA",
            "perihal": "Konfirmasi Pembayaran Pesanan",
            "isi": "Konfirmasi pelunasan sisa pembayaran pesanan Nomor 002/341/MS/X/2026 sebesar Rp17.375.000 melalui transfer bank tanggal 28 Oktober 2026.",
            "kodeArsip": "371",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 10,
            "noAgenda": 10,
            "tglTerima": "2026-11-02",
            "noSurat": "087/331/KMI/X/2026",
            "tglSurat": "2026-10-30",
            "pengirim": "PT KARYA MULTIPANEL INDONESIA",
            "perihal": "Penawaran Bahan Baku Furniture",
            "isi": "Penawaran papan MDF, HPL, plywood, rel laci, dan engsel hidrolik beserta syarat pembayaran; berlaku sampai 30 November 2026.",
            "kodeArsip": "331",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 11,
            "noAgenda": 11,
            "tglTerima": "2026-11-03",
            "noSurat": "015/312/SCE/XI/2026",
            "tglSurat": "2026-11-02",
            "pengirim": "PT SAMUDERA CARGO EKSPRESS",
            "perihal": "Perkenalan Jasa Pengiriman Barang",
            "isi": "Memperkenalkan layanan ekspedisi darat dan laut, pengepakan peti kayu, dan asuransi pengiriman serta menawarkan kerja sama.",
            "kodeArsip": "312",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 12,
            "noAgenda": 12,
            "tglTerima": "2026-11-03",
            "noSurat": "",
            "tglSurat": "2026-11-03",
            "pengirim": "Nurhaliza",
            "perihal": "Permohonan Cuti Izin",
            "isi": "Mengajukan cuti izin 3 (tiga) hari kerja, 9 s.d. 11 November 2026, untuk menghadiri pernikahan saudara kandung.",
            "kodeArsip": "113",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 13,
            "noAgenda": 13,
            "tglTerima": "2026-11-06",
            "noSurat": "026/341/PZ/XI/2026",
            "tglSurat": "2026-11-05",
            "pengirim": "PT PUTRI ZALSABILA",
            "perihal": "Pemesanan Barang",
            "isi": "Pesanan Meja Kerja Staff Minimalis 8 unit, Kursi Kantor Staff 8 unit, Meja Meeting/Rapat 2 unit, dan Rak Buku/Dokumen 4 unit senilai Rp34.900.000.",
            "kodeArsip": "341",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 14,
            "noAgenda": 14,
            "tglTerima": "2026-11-06",
            "noSurat": "",
            "tglSurat": "2026-11-06",
            "pengirim": "Dewi Anggraini",
            "perihal": "Permohonan Cuti Melahirkan",
            "isi": "Mengajukan cuti melahirkan 3 (tiga) bulan, 1 Desember 2026 s.d. 28 Februari 2027; perkiraan persalinan 15 Desember 2026.",
            "kodeArsip": "111",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 15,
            "noAgenda": 15,
            "tglTerima": "2026-11-10",
            "noSurat": "024/321/NTC/XI/2026",
            "tglSurat": "2026-11-09",
            "pengirim": "PT NAVION TECHNOLOGIES",
            "perihal": "Permintaan Penawaran Harga Furniture Kantor",
            "isi": "Memohon penawaran harga meja kerja staff, kursi ergonomis, meja meeting, dan lemari arsip untuk kantor cabang Makassar.",
            "kodeArsip": "321",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 16,
            "noAgenda": 16,
            "tglTerima": "2026-11-10",
            "noSurat": "",
            "tglSurat": "2026-11-10",
            "pengirim": "Andi Muhammad Fadli",
            "perihal": "Pengunduran Diri",
            "isi": "Mengajukan pengunduran diri sebagai Staf Produksi, efektif 10 Desember 2026, karena alasan pribadi.",
            "kodeArsip": "123",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 17,
            "noAgenda": 17,
            "tglTerima": "2026-11-12",
            "noSurat": "103/361/JP/XI/2026",
            "tglSurat": "2026-11-11",
            "pengirim": "CV JATI PRATAMA",
            "perihal": "Tagihan Pembayaran Bahan Baku Kayu",
            "isi": "Tagihan pengiriman kayu jati olahan dan papan jati sebesar Rp55.000.000, jatuh tempo 20 November 2026.",
            "kodeArsip": "361",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 18,
            "noAgenda": 18,
            "tglTerima": "2026-11-13",
            "noSurat": "",
            "tglSurat": "2026-11-12",
            "pengirim": "Nur Fadhilah",
            "perihal": "Lamaran Pekerjaan - Staf Desain Produk",
            "isi": "Mengajukan lamaran sebagai Staf Desain Produk; lulusan S1 Desain Interior dengan pengalaman 2 tahun.",
            "kodeArsip": "121",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 19,
            "noAgenda": 19,
            "tglTerima": "2026-11-13",
            "noSurat": "045/410/AIMS/XI/2026",
            "tglSurat": "2026-11-13",
            "pengirim": "ASOSIASI INDUSTRI MEBEL SULAWESI SELATAN",
            "perihal": "Undangan Rapat Anggota dan Seminar",
            "isi": "Mengundang menghadiri Rapat Anggota Tahunan dan Seminar Ekspor Mebel, Sabtu 28 November 2026 pukul 09.00 WITA di Makassar.",
            "kodeArsip": "410",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 20,
            "noAgenda": 20,
            "tglTerima": "2026-11-16",
            "noSurat": "",
            "tglSurat": "2026-11-16",
            "pengirim": "Wahyu Pratama",
            "perihal": "Permohonan Izin Mengikuti Pelatihan",
            "isi": "Memohon izin dan dukungan biaya (Rp3.500.000) mengikuti pelatihan digital marketing 7 s.d. 9 Desember 2026 di Makassar.",
            "kodeArsip": "142",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 21,
            "noAgenda": 21,
            "tglTerima": "2026-11-18",
            "noSurat": "112/480/SMKN2-MRS/XI/2026",
            "tglSurat": "2026-11-17",
            "pengirim": "SMK NEGERI 2 MAROS",
            "perihal": "Permohonan Praktik Kerja Lapangan (PKL)",
            "isi": "Memohon penerimaan 6 (enam) siswa PKL jurusan Desain dan Produksi Kria Kayu, 4 Januari s.d. 31 Maret 2027.",
            "kodeArsip": "480",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        },
        {
            "id": 22,
            "noAgenda": 22,
            "tglTerima": "2026-11-20",
            "noSurat": "",
            "tglSurat": "2026-11-20",
            "pengirim": "Muh. Aldi Saputra",
            "perihal": "Pengunduran Diri",
            "isi": "Mengajukan pengunduran diri sebagai Staf Logistik, efektif 31 Desember 2026, untuk melanjutkan pendidikan S1.",
            "kodeArsip": "123",
            "fileName": "",
            "fileId": null,
            "disposisi": null,
            "source": "BASM"
        }
    ],
    "suratKeluar": [
        {
            "id": 1,
            "noAgenda": 1,
            "tglKirim": "2026-09-02",
            "noSurat": "001/311/FRD/IX/2026",
            "tglSurat": "2026-09-01",
            "kepada": "PT PUTRI ZALSABILA",
            "perihal": "Perkenalan Perusahaan",
            "isi": "Memperkenalkan perusahaan PT Furnindo sebagai perusahaan yang bergerak di bidang furniture yang telah berdiri sejak 2017 dan telah memiliki beberapa cabang di wilayah Sulawesi.",
            "kode": "311",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 2,
            "noAgenda": 2,
            "tglKirim": "2026-09-10",
            "noSurat": "002/331/FRD/IX/2026",
            "tglSurat": "2026-09-10",
            "kepada": "PT PUTRI ZALSABILA",
            "perihal": "Balasan Permintaan Daftar Harga",
            "isi": "Balasan atas surat Nomor 004/121.1/PZ/IX/2026 tanggal 4 September 2026 perihal Permintaan Daftar Harga",
            "kode": "331",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 3,
            "noAgenda": 3,
            "tglKirim": "2026-09-11",
            "noSurat": "003/113/FRD/IX/2026",
            "tglSurat": "2026-09-10",
            "kepada": "Siti Aisyah",
            "perihal": "Persetujuan Cuti Menikah",
            "isi": "Menyetujui permohonan cuti menikah dari tanggal 15 September 2026 s/d 21 September 2026",
            "kode": "113",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 4,
            "noAgenda": 4,
            "tglKirim": "2026-09-11",
            "noSurat": "004/450/FRD/IX/2026",
            "tglSurat": "2026-09-11",
            "kepada": "Umum",
            "perihal": "Pengumuman Lowongan Kerja Manajer Pemasaran",
            "isi": "Membuka lowongan pekerjaan posisi Manajer Pemasaran dengan kualifikasi S1 Manajemen/Pemasaran/Bisnis, pengalaman minimal 3 tahun, Memiliki kemampuan leadership, analisis pasar, dan strategi pemasaran, Komunikatif, terbiasa bekerja dengan target, dan bersedia dinas ke cabang di wilayah Sulawesi.",
            "kode": "450",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 5,
            "noAgenda": 5,
            "tglKirim": "2026-09-15",
            "noSurat": "005/124/FRD/IX/2026",
            "tglSurat": "2026-09-15",
            "kepada": "Indriani",
            "perihal": "Panggilan Wawancara Kerja",
            "isi": "mengundang Saudara/i untuk mengikuti tahap wawancara kerja yang akan dilaksanakan pada Jumat, 18 September 2026; Pukul 09.00 WITA - selesai; di Kantor PT FURNINDO, Jl. Lingkungan Padangalla, Taroada, Kec. Mandai, Kabupaten Maros.",
            "kode": "124",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 6,
            "noAgenda": 6,
            "tglKirim": "2026-09-15",
            "noSurat": "006/124/FRD/IX/2026",
            "tglSurat": "2026-09-15",
            "kepada": "Wahyu Pratama",
            "perihal": "Panggilan Wawancara Kerja",
            "isi": "mengundang Saudara/i untuk mengikuti tahap wawancara kerja yang akan dilaksanakan pada Jumat, 18 September 2026; Pukul 09.00 WITA - selesai; di Kantor PT FURNINDO, Jl. Lingkungan Padangalla, Taroada, Kec. Mandai, Kabupaten Maros.",
            "kode": "124",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 7,
            "noAgenda": 7,
            "tglKirim": "2026-09-25",
            "noSurat": "007/131/FRD/IX/2026",
            "tglSurat": "2026-09-25",
            "kepada": "Wahyu Pratama",
            "perihal": "Pengangkatan Karyawan",
            "isi": "Mengangkat Saudara Wahyu Pratama sebagai Manajer Pemasaran PT   FURNINDO terhitung mulai tanggal 1 Oktober 2026.",
            "kode": "131",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 8,
            "noAgenda": 8,
            "tglKirim": "2026-10-02",
            "noSurat": "008/311/FRD/X/2026",
            "tglSurat": "2026-10-01",
            "kepada": "CV MITRA SEJAHTERA",
            "perihal": "Perkenalan Perusahaan",
            "isi": "Memperkenalkan Perusahaan PT FURNINDO sebagai perusahaan yang bergerak di bidang produksi dan distribusi furniture, baik untuk kebutuhan rumah tangga maupun perkantoran. Perusahaan kami telah berdiri sejak tahun 2017 dan telah memiliki beberapa cabang di wilayah Sulawesi.",
            "kode": "311",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 9,
            "noAgenda": 9,
            "tglKirim": "2026-10-10",
            "noSurat": "009/331/FRD/X/2026",
            "tglSurat": "2026-10-10",
            "kepada": "CV MITRA SEJAHTERA",
            "perihal": "Penawaran Barang",
            "isi": "Penawaran barang berupa Meja Kerja Staff Minimalis, Meja Kerja Manager, Kursi Kantor Staff, Kursi Kantor Manager, Lemari Arsip/Filing Cabinet 4 Pintu.",
            "kode": "331",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 10,
            "noAgenda": 10,
            "tglKirim": "2026-10-20",
            "noSurat": "010/352/FRD/X/2026",
            "tglSurat": "2026-10-19",
            "kepada": "CV MITRA SEJAHTERA",
            "perihal": "Pemberitahuan Pengiriman Pesanan",
            "isi": "Pemberitahuan pengiriman pesanan berupa Meja Kerja Staff Minimalis 10 unit; Kursi Kantor Staff 10 unit; Lemari Arsip/Filing Cabinet 4 Pintu 5 unit. Diperkirakan akan sampai sekitar 3-5 hari ke lokasi pemesan.",
            "kode": "352",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 11,
            "noAgenda": 11,
            "tglKirim": "2026-10-20",
            "noSurat": "011/353/FRD/X/2026",
            "tglSurat": "2026-10-19",
            "kepada": "CV MITRA SEJAHTERA",
            "perihal": "Surat Pengantar Barang",
            "isi": "Pengantar barang berupa Meja Kerja Staff Minimalis 10 unit; Kursi Kantor Staff 10 unit; Lemari Arsip/Filing Cabinet 4 Pintu 5 unit.",
            "kode": "353",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 12,
            "noAgenda": 12,
            "tglKirim": "2026-10-20",
            "noSurat": "012/361/FRD/X/2026",
            "tglSurat": "2026-10-19",
            "kepada": "CV MITRA SEJAHTERA",
            "perihal": "Faktur",
            "isi": "Faktur penjualan sebesar Rp34.750.000",
            "kode": "361",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 13,
            "noAgenda": 13,
            "tglKirim": "2026-10-20",
            "noSurat": "",
            "tglSurat": "2026-10-19",
            "kepada": "CV MITRA SEJAHTERA",
            "perihal": "Daftar Rincian Barang",
            "isi": "Rincian semua barang pesanan seperti Ukuran peti, Isi Peti, berat(netto & brutto).",
            "kode": "351",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 14,
            "noAgenda": 14,
            "tglKirim": "2026-10-20",
            "noSurat": "013/493/FRD/X/2026",
            "tglSurat": "2026-10-20",
            "kepada": "Wahyu Pratama",
            "perihal": "Perjanjian Kerja",
            "isi": "Perjanjian kerja atas nama Wahyu Pratama sebagai Manajer Pemasaran selama setahun.",
            "kode": "493",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 15,
            "noAgenda": 15,
            "tglKirim": "2026-10-28",
            "noSurat": "014/112/FRD/X/2026",
            "tglSurat": "2026-10-28",
            "kepada": "Rahmat Hidayat",
            "perihal": "Persetujuan Cuti Sakit",
            "isi": "Menyetujui permohonan cuti sakit selama 3 (tiga) hari kerja, 27 s.d. 29 Oktober 2026; kembali bekerja 30 Oktober 2026.",
            "kode": "112",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 16,
            "noAgenda": 16,
            "tglKirim": "2026-10-30",
            "noSurat": "015/391/FRD/X/2026",
            "tglSurat": "2026-10-30",
            "kepada": "CV MITRA SEJAHTERA",
            "perihal": "Balasan Konfirmasi Pembayaran",
            "isi": "Menerima pelunasan Rp17.375.000 (Faktur 012/361/FRD/X/2026 senilai Rp34.750.000 lunas) dan menindaklanjuti klaim 003/381/MS/X/2026: penggantian 2 unit meja dan 1 engsel laci rusak serta pengiriman kekurangan 1 unit lemari arsip.",
            "kode": "391",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 17,
            "noAgenda": 17,
            "tglKirim": "2026-11-09",
            "noSurat": "016/111/FRD/XI/2026",
            "tglSurat": "2026-11-09",
            "kepada": "Dewi Anggraini",
            "perihal": "Persetujuan Cuti Melahirkan",
            "isi": "Menyetujui cuti melahirkan 3 (tiga) bulan, 1 Desember 2026 s.d. 28 Februari 2027; kembali bekerja 1 Maret 2027.",
            "kode": "111",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 18,
            "noAgenda": 18,
            "tglKirim": "2026-11-12",
            "noSurat": "017/123/FRD/XI/2026",
            "tglSurat": "2026-11-12",
            "kepada": "Andi Muhammad Fadli",
            "perihal": "Persetujuan Pengunduran Diri",
            "isi": "Menyetujui pengunduran diri sebagai Staf Produksi, efektif 10 Desember 2026, disertai serah terima pekerjaan dan pengembalian inventaris.",
            "kode": "123",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 19,
            "noAgenda": 19,
            "tglKirim": "2026-11-19",
            "noSurat": "018/480/FRD/XI/2026",
            "tglSurat": "2026-11-19",
            "kepada": "SMK NEGERI 2 MAROS",
            "perihal": "Persetujuan Praktik Kerja Lapangan",
            "isi": "Bersedia menerima 6 (enam) siswa kelas XI jurusan Desain dan Produksi Kria Kayu untuk PKL, 4 Januari s.d. 31 Maret 2027.",
            "kode": "480",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 20,
            "noAgenda": 20,
            "tglKirim": "2026-11-20",
            "noSurat": "019/124/FRD/XI/2026",
            "tglSurat": "2026-11-20",
            "kepada": "Nur Fadhilah",
            "perihal": "Panggilan Wawancara Kerja",
            "isi": "Mengundang wawancara kerja posisi Staf Desain Produk, Jumat, 27 November 2026 pukul 09.00 WITA di Kantor PT FURNINDO, Kabupaten Maros.",
            "kode": "124",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        },
        {
            "id": 21,
            "noAgenda": 21,
            "tglKirim": "2026-11-23",
            "noSurat": "020/392/FRD/XI/2026",
            "tglSurat": "2026-11-23",
            "kepada": "ASOSIASI INDUSTRI MEBEL SULAWESI SELATAN",
            "perihal": "Konfirmasi Kehadiran Undangan",
            "isi": "Konfirmasi kehadiran pada Rapat Anggota Tahunan dan seminar Strategi Ekspor Mebel di Era Digital, Sabtu, 28 November 2026 pukul 09.00 WITA.",
            "kode": "392",
            "fileName": "",
            "fileId": null,
            "source": "BASK"
        }
    ],
    "klasifikasi": [
        {
            "utama": "100 KP Kepegawaian",
            "rincian": ""
        },
        {
            "utama": "    110 - CT  Cuti",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "111 - Cuti Melahirkan"
        },
        {
            "utama": "",
            "rincian": "112 - Cuti Sakit"
        },
        {
            "utama": "",
            "rincian": "113 - Cuti Izin"
        },
        {
            "utama": "    120 - PD  Pengadaan Pegawai",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "121 - Lamaran"
        },
        {
            "utama": "",
            "rincian": "122 - Pengangkatan"
        },
        {
            "utama": "",
            "rincian": "123 - Pemutusan Hubungan Kerja"
        },
        {
            "utama": "",
            "rincian": "124 - Wawancara"
        },
        {
            "utama": "    130 - SK  Surat Keputusan Pegawai",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "131 - Keterangan Kerja"
        },
        {
            "utama": "",
            "rincian": "132 - Promosi Jabatan"
        },
        {
            "utama": "    140 - MT  Mutasi",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "141 -  Alih Tugas"
        },
        {
            "utama": "",
            "rincian": "142 - Dinas Pendidikan / Pelatihan"
        },
        {
            "utama": "    150 - Pelatihan / Diklat",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "151 -  Internal"
        },
        {
            "utama": "",
            "rincian": "152 -  Eksternal"
        },
        {
            "utama": "    160 - SP  Surat Peringatan",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "161 - Peringatan Pertama"
        },
        {
            "utama": "",
            "rincian": "162 - Peringatan Kedua"
        },
        {
            "utama": "    170 - PN  Penilaian Kinerja",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "171 -  Hasil Review"
        },
        {
            "utama": "    180 - KR  Karyawan",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "181 - Tetap"
        },
        {
            "utama": "",
            "rincian": "182 -  Honorer / Kontrak"
        },
        {
            "utama": "    190 - CT  Catatan Pegawai",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "191 -  Absensi"
        },
        {
            "utama": "",
            "rincian": "192 -  Prestasi dan Disiplin"
        },
        {
            "utama": "200 KU Keuangan",
            "rincian": ""
        },
        {
            "utama": "    210 - GJ  Gaji",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "211 -  Tetap"
        },
        {
            "utama": "",
            "rincian": "212 -  Insentif"
        },
        {
            "utama": "",
            "rincian": "213 -  Tunjangan Jabatan"
        },
        {
            "utama": "    220 - BY  Biaya",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "221 -  Operasional"
        },
        {
            "utama": "",
            "rincian": "222 -  Transportasi"
        },
        {
            "utama": "",
            "rincian": "223 -  Akomodasi"
        },
        {
            "utama": "    230 - PD  Pendapatan",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "231 -  Uang Saku"
        },
        {
            "utama": "",
            "rincian": "232 -  Penghasilan Lain"
        },
        {
            "utama": "    240 - PT  Pertanggungjawaban",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "241 -  Laporan SPJ"
        },
        {
            "utama": "",
            "rincian": "242 -  Laporan Pertanggungjawaban Kegiatan"
        },
        {
            "utama": "    250 - PI  Piutang",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "251 -  Kolektibilitas"
        },
        {
            "utama": "    260 - HU  Hutang",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "261 -  Supplier"
        },
        {
            "utama": "",
            "rincian": "262 -  Pribadi"
        },
        {
            "utama": "    270 - LK  Laporan Keuangan",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "271 -  Bulanan"
        },
        {
            "utama": "",
            "rincian": "272 -  Tahunan"
        },
        {
            "utama": "    280 - BK  Bukti Kas",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "281 -  Masuk"
        },
        {
            "utama": "",
            "rincian": "282 -  Keluar"
        },
        {
            "utama": "    290 - AU  Audit",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "291 -  Internal"
        },
        {
            "utama": "",
            "rincian": "292 -  Eksternal"
        },
        {
            "utama": "300 BS Bisnis / Niaga",
            "rincian": ""
        },
        {
            "utama": "    310 - PK  Perkenalan",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "311 -  Perkenalan Perusahaan"
        },
        {
            "utama": "",
            "rincian": "312 -  Perkenalan Produk / Jasa"
        },
        {
            "utama": "    320 - PP  Permintaan Penawaran",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "321 -  Harga Barang"
        },
        {
            "utama": "",
            "rincian": "322 -  Jasa / Layanan"
        },
        {
            "utama": "    330 - PN  Penawaran",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "331 -  Barang"
        },
        {
            "utama": "",
            "rincian": "332 -  Jasa"
        },
        {
            "utama": "    340 - PS  Pemesanan",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "341 -  Barang"
        },
        {
            "utama": "",
            "rincian": "342 -  Jasa"
        },
        {
            "utama": "    350 - PG  Pengiriman Barang",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "351 -  Surat Jalan / Pengantar"
        },
        {
            "utama": "",
            "rincian": "352 -  Pemberitahuan Pengiriman"
        },
        {
            "utama": "",
            "rincian": "353 - Pengantaran Barang"
        },
        {
            "utama": "    360 - FT  Faktur / Tagihan",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "361 -  Tagihan Pembayaran"
        },
        {
            "utama": "",
            "rincian": "362 -  Faktur Pajak"
        },
        {
            "utama": "    370 - PM  Pembayaran",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "371 -  Lunas / Konfirmasi Pembayaran"
        },
        {
            "utama": "    380 - KL  Klaim",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "381 -  Barang Rusak / Tidak Sesuai"
        },
        {
            "utama": "",
            "rincian": "382 -  Keterlambatan Servis"
        },
        {
            "utama": "    390 - BK  Balasan / Konfirmasi",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "391 -  Surat Balasan Bisnis"
        },
        {
            "utama": "",
            "rincian": "392 -  Konfirmasi Surat Sebelumnya"
        },
        {
            "utama": "410 – UD  Surat Undangan",
            "rincian": ""
        },
        {
            "utama": "  420 – KU  Surat Kuasa",
            "rincian": ""
        },
        {
            "utama": "  430 – KT  Surat Keterangan",
            "rincian": ""
        },
        {
            "utama": "  440 – ND  Nota Dinas",
            "rincian": ""
        },
        {
            "utama": "  450 – PM  Surat Pengumuman",
            "rincian": ""
        },
        {
            "utama": "  460 – ED  Surat Edaran",
            "rincian": ""
        },
        {
            "utama": "  470 – PG  Surat Pengantar",
            "rincian": ""
        },
        {
            "utama": "  480 – PH  Surat Permohonan",
            "rincian": ""
        },
        {
            "utama": "  490 – SK  Surat Keputusan",
            "rincian": ""
        },
        {
            "utama": "",
            "rincian": "  491 – TG  Surat Tugas"
        },
        {
            "utama": "",
            "rincian": "  492 – PR  Surat Perintah"
        },
        {
            "utama": "",
            "rincian": "  493 – PJ   Surat Perjanjian"
        },
        {
            "utama": "",
            "rincian": "  494 – BA Berita Acara"
        }
    ]
};

const STORAGE_KEYS = {
    masuk: 'hmapd_surat_masuk_v4',
    keluar: 'hmapd_surat_keluar_v4',
    klasifikasi: 'hmapd_klasifikasi_v4',
    kontakManajer: 'hmapd_kontak_manajer_v1',
    kontakAdmin: 'hmapd_kontak_admin_v1',
    tema: 'hmapd_theme'
};

let suratMasukData = loadStoredRecords(STORAGE_KEYS.masuk, DATA_EXCEL.suratMasuk, 'masuk');
let suratKeluarData = loadStoredRecords(STORAGE_KEYS.keluar, DATA_EXCEL.suratKeluar, 'keluar');
let klasifikasiData = loadStoredClassification();
let kontakManajerData = loadStoredKontakManajer();
let kontakAdminData = loadStoredKontakAdmin();
let currentRole = 'admin';
let selectedSuratId = null;
let previewObjectUrl = null;
let fileDbPromise = null;

/* ----------------------- Penyimpanan data ----------------------- */
function deepCopy(value) {
    return JSON.parse(JSON.stringify(value));
}

function safeStorageGet(key) {
    try {
        return localStorage.getItem(key);
    } catch (error) {
        console.error('LocalStorage tidak dapat dibaca:', error);
        return null;
    }
}

function safeStorageSet(key, value) {
    try {
        localStorage.setItem(key, value);
        return true;
    } catch (error) {
        console.error('Gagal menyimpan data:', error);
        alert('Penyimpanan browser penuh atau tidak tersedia. Data belum berhasil disimpan.');
        return false;
    }
}

function parseStoredArray(key) {
    const raw = safeStorageGet(key);
    if (raw === null) return null;
    try {
        const parsed = JSON.parse(raw);
        return Array.isArray(parsed) ? parsed : null;
    } catch (error) {
        // Simpan salinan nilai rusak untuk mencegah data lama hilang tanpa jejak.
        safeStorageSet(`hmapd_backup_corrupt_${key}`, raw);
        console.warn(`Data ${key} tidak valid; menggunakan data Excel.`, error);
        return null;
    }
}

function normaliseIdentity(value) {
    return String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function sameArchiveRecord(a, b, type) {
    const aNo = normaliseIdentity(a.noSurat);
    const bNo = normaliseIdentity(b.noSurat);
    if (aNo && bNo) return aNo === bNo;
    const aDate = type === 'masuk' ? a.tglTerima : a.tglKirim;
    const bDate = type === 'masuk' ? b.tglTerima : b.tglKirim;
    const aParty = type === 'masuk' ? a.pengirim : a.kepada;
    const bParty = type === 'masuk' ? b.pengirim : b.kepada;
    return Boolean(aDate && bDate && normaliseIdentity(aDate) === normaliseIdentity(bDate)
        && normaliseIdentity(aParty) === normaliseIdentity(bParty));
}

function legacyDisposition(record) {
    if (record.disposisi && typeof record.disposisi === 'object') return record.disposisi;
    if (!record.instruksi && !record.tglSelesai && !record.penerima) return null;
    return {
        tanggalSelesai: record.tglSelesai || '',
        penerima: record.penerima || '',
        tindakan: record.instruksi || '',
        prioritas: record.prioritas || '',
        lampiran: record.lampiran || '',
        status: record.status === 'Selesai' ? 'Selesai' : 'Dalam Proses'
    };
}

function loadStoredRecords(key, excelDefaults, type) {
    const stored = parseStoredArray(key);
    if (!stored || stored.length === 0) return deepCopy(excelDefaults);

    // Jadikan baris Excel sebagai dasar; pindahkan lampiran/disposisi lama bila cocok.
    const merged = deepCopy(excelDefaults);
    const usedIds = new Set(merged.map(item => String(item.id)));
    let nextId = Math.max(0, ...merged.map(item => Number(item.id) || 0)) + 1;

    stored.forEach(oldRecord => {
        if (!oldRecord || typeof oldRecord !== 'object') return;
        const match = merged.find(item => sameArchiveRecord(item, oldRecord, type));
        if (match) {
            if (oldRecord.fileId) {
                match.fileId = oldRecord.fileId;
                match.fileName = oldRecord.fileName || match.fileName;
            }
            const prior = legacyDisposition(oldRecord);
            if (prior) match.disposisi = prior;
            return;
        }
        const preserved = { ...oldRecord };
        let id = Number(preserved.id);
        if (!Number.isFinite(id) || usedIds.has(String(id))) id = nextId++;
        while (usedIds.has(String(id))) id = nextId++;
        preserved.id = id;
        usedIds.add(String(id));
        if (type === 'masuk' && !preserved.disposisi) preserved.disposisi = legacyDisposition(preserved);
        merged.push(preserved);
    });
    return merged;
}

function loadStoredClassification() {
    const stored = parseStoredArray(STORAGE_KEYS.klasifikasi);
    if (!stored || stored.length === 0) return deepCopy(DATA_EXCEL.klasifikasi);
    // Format lama berisi empat ringkasan string, bukan tabel hierarki Excel.
    if (stored.every(item => item && typeof item === 'object' && ('utama' in item || 'rincian' in item))) {
        return stored;
    }
    safeStorageSet('hmapd_backup_klasifikasi_lama', JSON.stringify(stored));
    return deepCopy(DATA_EXCEL.klasifikasi);
}

function loadStoredKontakManajer() {
    const stored = parseStoredArray(STORAGE_KEYS.kontakManajer);
    if (!stored || stored.length === 0) return deepCopy(DEFAULT_KONTAK_MANAJER);
    return stored;
}

function loadStoredKontakAdmin() {
    const raw = safeStorageGet(STORAGE_KEYS.kontakAdmin);
    if (!raw) return deepCopy(DEFAULT_KONTAK_ADMIN);
    try {
        const parsed = JSON.parse(raw);
        return parsed && typeof parsed === 'object' ? parsed : deepCopy(DEFAULT_KONTAK_ADMIN);
    } catch (_) {
        return deepCopy(DEFAULT_KONTAK_ADMIN);
    }
}

function saveDataToStorage() {
    const ok1 = safeStorageSet(STORAGE_KEYS.masuk, JSON.stringify(suratMasukData));
    const ok2 = safeStorageSet(STORAGE_KEYS.keluar, JSON.stringify(suratKeluarData));
    const ok3 = safeStorageSet(STORAGE_KEYS.klasifikasi, JSON.stringify(klasifikasiData));
    const ok4 = safeStorageSet(STORAGE_KEYS.kontakManajer, JSON.stringify(kontakManajerData));
    const ok5 = safeStorageSet(STORAGE_KEYS.kontakAdmin, JSON.stringify(kontakAdminData));
    pushAllToSupabase();
    return ok1 && ok2 && ok3 && ok4 && ok5;
}

/* ----------------------- Integrasi Supabase Cloud ----------------------- */
function initSupabase() {
    const badge = document.getElementById('supabaseStatusBadge');
    const desc = document.getElementById('supabaseStatusDesc');
    if (SUPABASE_URL && SUPABASE_ANON_KEY && window.supabase?.createClient) {
        try {
            supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
            if (badge) {
                badge.textContent = 'Terhubung ke Supabase';
                badge.className = 'badge bg-success';
            }
            if (desc) desc.textContent = 'Data otomatis tersinkronisasi ke database cloud Supabase untuk semua perangkat.';
            return true;
        } catch (err) {
            console.error('Gagal inisialisasi Supabase:', err);
        }
    }
    if (badge) {
        badge.textContent = 'Mode Lokal';
        badge.className = 'badge bg-secondary';
    }
    return false;
}

async function pushAllToSupabase() {
    if (!supabaseClient) return;
    const rows = [
        { key: 'surat_masuk', data: suratMasukData, updated_at: new Date().toISOString() },
        { key: 'surat_keluar', data: suratKeluarData, updated_at: new Date().toISOString() },
        { key: 'klasifikasi', data: klasifikasiData, updated_at: new Date().toISOString() },
        { key: 'kontak_manajer', data: kontakManajerData, updated_at: new Date().toISOString() },
        { key: 'kontak_admin', data: kontakAdminData, updated_at: new Date().toISOString() }
    ];
    try {
        const { error } = await supabaseClient.from('hmapd_store').upsert(rows, { onConflict: 'key' });
        if (error) console.error('Supabase upsert error:', error);
    } catch (err) {
        console.error('Supabase push error:', err);
    }
}

async function syncFromSupabase(showNotice = false) {
    if (!supabaseClient) {
        if (showNotice) alert('Supabase belum dikonfigurasi. Isi SUPABASE_URL dan SUPABASE_ANON_KEY di file script.js terlebih dahulu.');
        return;
    }
    try {
        const { data, error } = await supabaseClient.from('hmapd_store').select('key, data');
        if (error) throw error;
        if (!data || data.length === 0) {
            await pushAllToSupabase();
            if (showNotice) alert('Data awal berhasil diunggah ke Supabase.');
            return;
        }
        const map = {};
        data.forEach(row => { map[row.key] = row.data; });
        if (Array.isArray(map.surat_masuk) && map.surat_masuk.length) suratMasukData = map.surat_masuk;
        if (Array.isArray(map.surat_keluar) && map.surat_keluar.length) suratKeluarData = map.surat_keluar;
        if (Array.isArray(map.klasifikasi) && map.klasifikasi.length) klasifikasiData = map.klasifikasi;
        if (Array.isArray(map.kontak_manajer) && map.kontak_manajer.length) kontakManajerData = map.kontak_manajer;
        if (map.kontak_admin && typeof map.kontak_admin === 'object') kontakAdminData = map.kontak_admin;

        safeStorageSet(STORAGE_KEYS.masuk, JSON.stringify(suratMasukData));
        safeStorageSet(STORAGE_KEYS.keluar, JSON.stringify(suratKeluarData));
        safeStorageSet(STORAGE_KEYS.klasifikasi, JSON.stringify(klasifikasiData));
        safeStorageSet(STORAGE_KEYS.kontakManajer, JSON.stringify(kontakManajerData));
        safeStorageSet(STORAGE_KEYS.kontakAdmin, JSON.stringify(kontakAdminData));
        renderAll();
        if (showNotice) alert('Data berhasil disinkronkan dari Supabase!');
    } catch (err) {
        console.error('Supabase sync error:', err);
        if (showNotice) alert('Gagal mengambil data dari Supabase. Periksa koneksi internet atau konfigurasi tabel.');
    }
}

/* ----------------------- Utilitas tampilan ----------------------- */
function escapeHTML(value) {
    return String(value ?? '').replace(/[&<>"']/g, char => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[char]));
}

function formatTanggal(value) {
    if (!value) return '-';
    const parsed = new Date(`${value}T00:00:00`);
    return Number.isNaN(parsed.getTime()) ? escapeHTML(value) : parsed.toLocaleDateString('id-ID', {
        day: '2-digit', month: 'short', year: 'numeric'
    });
}

function formatTanggalDisposisi(value) {
    if (!value) return '-';
    const parsed = new Date(`${value}T00:00:00`);
    return Number.isNaN(parsed.getTime()) ? escapeHTML(value) : parsed.toLocaleDateString('id-ID', {
        day: '2-digit', month: 'long', year: 'numeric'
    });
}

function textOrDash(value) {
    return value === null || value === undefined || String(value).trim() === '' ? '-' : escapeHTML(value);
}

function getModal(id) {
    const element = document.getElementById(id);
    if (!element || !window.bootstrap?.Modal) return null;
    return bootstrap.Modal.getOrCreateInstance
        ? bootstrap.Modal.getOrCreateInstance(element)
        : new bootstrap.Modal(element);
}

/* ----------------------- Login dan hak akses ----------------------- */
function togglePasswordInput() {
    const role = document.getElementById('roleSelect')?.value || 'admin';
    const group = document.getElementById('passwordGroup');
    const label = document.getElementById('passwordLabel');
    const password = document.getElementById('accessPassword');
    if (!group || !password) return;
    const needsPassword = role !== 'tamu';
    group.style.display = needsPassword ? 'block' : 'none';
    password.required = needsPassword;
    if (label) label.textContent = role === 'pimpinan' ? 'Password Pimpinan' : 'Password Admin';
    if (!needsPassword) password.value = '';
}

function handleLogin(event) {
    event?.preventDefault();
    const role = document.getElementById('roleSelect')?.value || 'admin';
    const pass = document.getElementById('accessPassword')?.value || '';
    if (role === 'admin' && pass !== '123456') {
        alert('Password Admin salah.');
        return false;
    }
    if (role === 'pimpinan' && pass !== '234567') {
        alert('Password Pimpinan salah.');
        return false;
    }

    currentRole = role;
    document.getElementById('authContainer')?.classList.add('d-none');
    document.getElementById('appContainer')?.classList.remove('d-none');
    const badge = document.getElementById('userRoleBadge');
    if (badge) badge.textContent = `Role: ${role.toUpperCase()}`;
    applyRolePermissions();
    renderAll();
    switchTab('beranda');
    return false;
}

function applyRolePermissions() {
    document.querySelectorAll('.admin-only').forEach(element => {
        element.classList.toggle('d-none', currentRole !== 'admin');
    });
    document.querySelectorAll('.pimpinan-only').forEach(element => {
        element.classList.toggle('d-none', currentRole !== 'pimpinan');
    });
    document.querySelectorAll('.classification-action-col').forEach(el => el.classList.toggle('d-none', currentRole !== 'admin'));
    const editor = document.getElementById('disposisiEditorContainer');
    if (editor) editor.classList.toggle('d-none', currentRole !== 'pimpinan');
    const admEditor = document.getElementById('adminDisposisiEditorContainer');
    if (admEditor) admEditor.classList.toggle('d-none', currentRole !== 'admin');
    const downloadButton = document.getElementById('btnDownloadDisposisi');
    if (downloadButton) downloadButton.classList.toggle('d-none', currentRole === 'tamu');
    const printButton = document.getElementById('btnPrintDisposisi');
    if (printButton) printButton.classList.toggle('d-none', currentRole === 'tamu');
}

function logout() {
    if (!confirm('Apakah Anda yakin ingin keluar dari Portal Kearsipan PT HMAPD?')) return false;
    selectedSuratId = null;
    document.getElementById('appContainer')?.classList.add('d-none');
    document.getElementById('authContainer')?.classList.remove('d-none');
    const passInput = document.getElementById('accessPassword');
    if (passInput) passInput.value = '';
    return false;
}

/* ----------------------- Navigasi menu ----------------------- */
function switchTab(tabId, clickedElement = null) {
    document.querySelectorAll('.tab-content-section').forEach(section => section.classList.add('d-none'));
    const target = document.getElementById(`tab-${tabId}`);
    if (!target) return false;
    target.classList.remove('d-none');

    document.querySelectorAll('.sidebar .nav-link').forEach(link => link.classList.remove('active'));
    if (clickedElement?.classList) clickedElement.classList.add('active');
    else {
        const activeLink = document.querySelector(`.sidebar .nav-link[onclick*="'${tabId}'"]`);
        activeLink?.classList.add('active');
    }

    const titles = {
        beranda: 'Beranda Utama',
        suratMasuk: 'Surat Masuk',
        suratKeluar: 'Surat Keluar',
        klasifikasi: 'Kode Klasifikasi',
        lembarDisposisi: 'Lembar Disposisi',
        pengaturanKontak: 'Pengaturan Kontak & Tujuan Disposisi'
    };
    const title = document.getElementById('currentTitle');
    if (title) title.textContent = titles[tabId] || 'Portal Kearsipan';

    if (tabId === 'suratMasuk') filterSuratMasuk();
    if (tabId === 'suratKeluar') filterSuratKeluar();
    if (tabId === 'klasifikasi') renderKlasifikasiTable();
    if (tabId === 'lembarDisposisi') renderDaftarDisposisi();
    if (tabId === 'pengaturanKontak') renderKontakSettings();
    document.getElementById('sidebar')?.classList.remove('mobile-open');
    return false;
}

function toggleMobileSidebar() {
    document.getElementById('sidebar')?.classList.toggle('mobile-open');
}

function navigateHomeSection(sectionId) {
    switchTab('beranda');
    window.setTimeout(() => {
        document.getElementById(sectionId)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 50);
    return false;
}

function toggleDropdown(event) {
    event?.preventDefault();
    const dropdown = document.getElementById('dropdownLayanan');
    dropdown?.classList.toggle('active');
    return false;
}

function handleSubMenu(menu) {
    document.getElementById('dropdownLayanan')?.classList.remove('active');
    if (menu === 'arsip') {
        switchTab('suratMasuk');
        return false;
    }
    if (menu === 'onboarding') {
        navigateHomeSection('profil-section');
        alert('Informasi onboarding perusahaan dapat dilihat pada bagian Profil & Informasi Perusahaan.');
        return false;
    }
    return false;
}

function toggleTheme() {
    document.body.classList.toggle('dark-mode');
    const dark = document.body.classList.contains('dark-mode');
    const icon = document.getElementById('themeIcon');
    if (icon) icon.className = dark ? 'fa-solid fa-sun' : 'fa-solid fa-moon';
    try { localStorage.setItem(STORAGE_KEYS.tema, dark ? 'dark' : 'light'); } catch (_) {}
}

function handleFeatureClick(feature) {
    const messages = {
        ecommerce: 'PT HMAPD menyediakan layanan perdagangan dan distribusi produk digital.',
        transaksi: 'Informasi ini menjelaskan komitmen perusahaan terhadap proses transaksi yang cepat.',
        lisensi: 'PT HMAPD mengutamakan produk digital dengan lisensi resmi dan aman.'
    };
    alert(messages[feature] || 'Informasi fitur PT HMAPD.');
}

/* ----------------------- Render tabel ----------------------- */
function renderAll() {
    renderSuratMasukTable(suratMasukData);
    renderSuratKeluarTable(suratKeluarData);
    renderKlasifikasiTable();
    renderDaftarDisposisi();
    renderKontakSettings();
    populateDisposisiPenerimaOptions();
}

function fileActionButtons(type, item) {
    if (!item.fileId) return '<span class="small text-muted">Belum ada lampiran</span>';
    return `
        <button type="button" class="btn btn-sm btn-info text-white mb-1" title="Preview file" onclick="previewFile('${type}', ${Number(item.id)})"><i class="fa-solid fa-eye"></i></button>
        <button type="button" class="btn btn-sm btn-success text-white mb-1" title="Unduh file" onclick="downloadFile('${type}', ${Number(item.id)})"><i class="fa-solid fa-download"></i></button>
    `;
}

function renderSuratMasukTable(dataList) {
    const tbody = document.getElementById('tbodySuratMasuk');
    if (!tbody) return;
    tbody.innerHTML = dataList.length ? dataList.map(item => {
        const actions = currentRole === 'admin' ? `
            <button type="button" class="btn btn-sm btn-outline-warning mb-1" onclick="bukaModalEditMasuk(${Number(item.id)})"><i class="fa-solid fa-pen-to-square"></i> Edit</button>
            <button type="button" class="btn btn-sm btn-outline-danger mb-1" onclick="hapusSuratMasuk(${Number(item.id)})"><i class="fa-solid fa-trash"></i> Hapus</button>
        ` : '';
        return `
            <tr>
                <td>${textOrDash(item.noAgenda)}</td>
                <td>${formatTanggal(item.tglTerima)}</td>
                <td class="fw-semibold">${textOrDash(item.noSurat)}</td>
                <td>${formatTanggal(item.tglSurat)}</td>
                <td>${textOrDash(item.pengirim)}</td>
                <td class="fw-semibold text-primary">${textOrDash(item.perihal)}</td>
                <td class="small cell-longtext">${textOrDash(item.isi)}</td>
                <td><span class="badge bg-secondary-subtle text-dark">${textOrDash(item.kodeArsip)}</span></td>
                <td>${fileActionButtons('masuk', item)}</td>
                <td>
                    ${actions || '<span class="text-muted">-</span>'}
                </td>
            </tr>`;
    }).join('') : '<tr><td colspan="10" class="text-center text-muted py-4">Belum ada data surat masuk.</td></tr>';
}

function renderSuratKeluarTable(dataList) {
    const tbody = document.getElementById('tbodySuratKeluar');
    if (!tbody) return;
    tbody.innerHTML = dataList.length ? dataList.map(item => {
        const actions = currentRole === 'admin' ? `
            <button type="button" class="btn btn-sm btn-outline-warning mb-1" onclick="bukaModalEditKeluar(${Number(item.id)})"><i class="fa-solid fa-pen-to-square"></i> Edit</button>
            <button type="button" class="btn btn-sm btn-outline-danger mb-1" onclick="hapusSuratKeluar(${Number(item.id)})"><i class="fa-solid fa-trash"></i> Hapus</button>
        ` : '';
        return `
            <tr>
                <td>${textOrDash(item.noAgenda)}</td>
                <td>${formatTanggal(item.tglKirim)}</td>
                <td class="fw-semibold">${textOrDash(item.noSurat)}</td>
                <td>${formatTanggal(item.tglSurat)}</td>
                <td>${textOrDash(item.kepada)}</td>
                <td class="fw-semibold text-primary">${textOrDash(item.perihal)}</td>
                <td class="small cell-longtext">${textOrDash(item.isi)}</td>
                <td><span class="badge bg-secondary-subtle text-dark">${textOrDash(item.kode)}</span></td>
                <td>${fileActionButtons('keluar', item)}</td>
                <td>${actions}</td>
            </tr>`;
    }).join('') : '<tr><td colspan="10" class="text-center text-muted py-4">Belum ada data surat keluar.</td></tr>';
}

function renderKlasifikasiTable(dataList = klasifikasiData) {
    const tbody = document.getElementById('tbodyKlasifikasi');
    if (!tbody) return;
    const rows = dataList.map(row => {
        const originalIndex = klasifikasiData.indexOf(row);
        const actions = currentRole === 'admin' ? `
            <button type="button" class="btn btn-sm btn-outline-warning mb-1" onclick="bukaModalEditKlasifikasi(${originalIndex})" title="Edit klasifikasi"><i class="fa-solid fa-pen-to-square"></i> Edit</button>
            <button type="button" class="btn btn-sm btn-outline-danger mb-1" onclick="hapusKlasifikasi(${originalIndex})" title="Hapus klasifikasi"><i class="fa-solid fa-trash"></i> Hapus</button>` : '<span class="text-muted">-</span>';
        return `<tr>
            <td class="classification-source-cell">${escapeHTML(row.utama || '')}</td>
            <td class="classification-source-cell">${escapeHTML(row.rincian || '')}</td>
            <td class="classification-action-col">${actions}</td>
        </tr>`;
    });
    tbody.innerHTML = rows.length ? rows.join('') : '<tr><td colspan="3" class="text-center text-muted">Belum ada data klasifikasi yang sesuai.</td></tr>';
}

function bukaModalTambahKlasifikasi() {
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat menambah klasifikasi.');
    document.getElementById('modalKlasifikasiTitle').textContent = 'Tambah Klasifikasi';
    document.getElementById('editKlasifikasiIndex').value = '';
    document.getElementById('formKlasifikasi')?.reset();
    getModal('formKlasifikasiModal')?.show();
}

function bukaModalEditKlasifikasi(index) {
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat mengedit klasifikasi.');
    const row = klasifikasiData[index];
    if (!row) return;
    document.getElementById('modalKlasifikasiTitle').textContent = 'Edit Klasifikasi';
    document.getElementById('editKlasifikasiIndex').value = String(index);
    document.getElementById('klasifikasiUtamaInput').value = row.utama || '';
    document.getElementById('klasifikasiRincianInput').value = row.rincian || '';
    getModal('formKlasifikasiModal')?.show();
}

function simpanKlasifikasi(event) {
    event?.preventDefault();
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat menyimpan klasifikasi.');
    const utama = document.getElementById('klasifikasiUtamaInput').value;
    const rincian = document.getElementById('klasifikasiRincianInput').value;
    if (!utama.trim() && !rincian.trim()) return alert('Isi kode/kelompok atau rincian klasifikasi.');
    const indexValue = document.getElementById('editKlasifikasiIndex').value;
    const row = { utama, rincian };
    if (indexValue === '') klasifikasiData.push(row);
    else {
        const index = Number(indexValue);
        if (!Number.isInteger(index) || !klasifikasiData[index]) return alert('Data klasifikasi yang diedit tidak ditemukan.');
        klasifikasiData[index] = row;
    }
    if (!saveDataToStorage()) return;
    const keyword = document.getElementById('searchKlasifikasi')?.value || '';
    const modal = bootstrap.Modal.getInstance(document.getElementById('formKlasifikasiModal'));
    modal?.hide();
    renderKlasifikasiTable(klasifikasiData.filter(item => normalizeSearchText(collectSearchableText(item)).includes(normalizeSearchText(keyword))));
}

function hapusKlasifikasi(index) {
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat menghapus klasifikasi.');
    if (!klasifikasiData[index]) return;
    if (!confirm('Hapus baris klasifikasi ini?')) return;
    klasifikasiData.splice(index, 1);
    if (!saveDataToStorage()) return;
    filterKlasifikasi();
}

/* Utilitas pencarian bersama: dipakai oleh surat masuk, surat keluar,
   klasifikasi, dan daftar disposisi. */
function normalizeSearchText(value) {
    return String(value ?? '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLocaleLowerCase('id-ID')
        .replace(/\s+/g, ' ')
        .trim();
}

function collectSearchableText(value) {
    if (value === null || value === undefined) return '';
    if (Array.isArray(value)) return value.map(collectSearchableText).join(' ');
    if (typeof value === 'object') return Object.values(value).map(collectSearchableText).join(' ');
    return String(value);
}

function filterSuratMasuk() {
    const keyword = normalizeSearchText(document.getElementById('searchMasuk')?.value || '');
    const filtered = suratMasukData.filter(item => normalizeSearchText(collectSearchableText(item)).includes(keyword));
    renderSuratMasukTable(filtered);
}

function filterSuratKeluar() {
    const keyword = normalizeSearchText(document.getElementById('searchKeluar')?.value || '');
    const filtered = suratKeluarData.filter(item => normalizeSearchText(collectSearchableText(item)).includes(keyword));
    renderSuratKeluarTable(filtered);
}

function filterKlasifikasi() {
    const keyword = normalizeSearchText(document.getElementById('searchKlasifikasi')?.value || '');
    const filtered = klasifikasiData.filter(item => normalizeSearchText(collectSearchableText(item)).includes(keyword));
    renderKlasifikasiTable(filtered);
}

/* ----------------------- Modal surat masuk/keluar ----------------------- */
function nextRecordId(records) {
    return Math.max(0, ...records.map(item => Number(item.id) || 0)) + 1;
}

function nextAgendaNumber(records) {
    return Math.max(0, ...records.map(item => Number(item.noAgenda) || 0)) + 1;
}

function bukaModalTambahMasuk() {
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat menambahkan surat masuk.');
    document.getElementById('modalMasukTitle').textContent = 'Tambah Surat Masuk';
    document.getElementById('editMasukId').value = '';
    document.getElementById('formInputMasuk')?.reset();
    getModal('formSuratMasukModal')?.show();
}

function bukaModalEditMasuk(id) {
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat mengedit surat masuk.');
    const record = suratMasukData.find(item => String(item.id) === String(id));
    if (!record) return;
    document.getElementById('modalMasukTitle').textContent = 'Edit Surat Masuk';
    document.getElementById('editMasukId').value = record.id;
    document.getElementById('inNoAgendaMasuk').value = record.noAgenda ?? '';
    document.getElementById('inTglMasuk').value = record.tglTerima || '';
    document.getElementById('inTglSuratMasuk').value = record.tglSurat || '';
    document.getElementById('inNoMasuk').value = record.noSurat || '';
    document.getElementById('inPengirim').value = record.pengirim || '';
    document.getElementById('inPerihalMasuk').value = record.perihal || '';
    document.getElementById('inIsiMasuk').value = record.isi || '';
    const elLamp = document.getElementById('inLampiranMasuk');
    if (elLamp) elLamp.value = record.lampiran || '';
    const elSifat = document.getElementById('inSifatMasuk');
    if (elSifat) elSifat.value = record.sifatSurat || record.disposisi?.prioritas || '';
    document.getElementById('inKodeArsipMasuk').value = record.kodeArsip || '';
    const file = document.getElementById('inFileMasuk');
    if (file) file.value = '';
    getModal('formSuratMasukModal')?.show();
}

async function simpanSuratMasuk(event) {
    event?.preventDefault();
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat menyimpan surat masuk.');
    const id = document.getElementById('editMasukId').value;
    const current = id ? suratMasukData.find(item => String(item.id) === String(id)) : null;
    const file = document.getElementById('inFileMasuk')?.files?.[0] || null;
    let fileId = current?.fileId || null;
    let fileName = current?.fileName || '';
    if (file) {
        try {
            fileId = await saveUploadedFile(file);
            fileName = file.name;
        } catch (error) {
            console.error(error);
            return alert('File belum berhasil disimpan pada penyimpanan browser.');
        }
    }
    const values = {
        noAgenda: document.getElementById('inNoAgendaMasuk').value === '' ? (current?.noAgenda ?? nextAgendaNumber(suratMasukData)) : Number(document.getElementById('inNoAgendaMasuk').value),
        tglTerima: document.getElementById('inTglMasuk').value,
        tglSurat: document.getElementById('inTglSuratMasuk').value,
        noSurat: document.getElementById('inNoMasuk').value.trim(),
        pengirim: document.getElementById('inPengirim').value.trim(),
        perihal: document.getElementById('inPerihalMasuk').value.trim(),
        isi: document.getElementById('inIsiMasuk').value.trim(),
        lampiran: (document.getElementById('inLampiranMasuk')?.value || '').trim(),
        sifatSurat: (document.getElementById('inSifatMasuk')?.value || '').trim(),
        kodeArsip: document.getElementById('inKodeArsipMasuk').value.trim(),
        fileId,
        fileName,
        source: current?.source || 'Input Manual'
    };
    if (current) Object.assign(current, values);
    else suratMasukData.push({ id: nextRecordId(suratMasukData), ...values, disposisi: null });
    if (!saveDataToStorage()) return;
    renderAll();
    getModal('formSuratMasukModal')?.hide();
    document.getElementById('formInputMasuk')?.reset();
    alert(current ? 'Surat masuk berhasil diperbarui.' : 'Surat masuk berhasil disimpan.');
}

function bukaModalTambahKeluar() {
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat menambahkan surat keluar.');
    document.getElementById('modalKeluarTitle').textContent = 'Tambah Surat Keluar';
    document.getElementById('editKeluarId').value = '';
    document.getElementById('formInputKeluar')?.reset();
    getModal('formSuratKeluarModal')?.show();
}

function bukaModalEditKeluar(id) {
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat mengedit surat keluar.');
    const record = suratKeluarData.find(item => String(item.id) === String(id));
    if (!record) return;
    document.getElementById('modalKeluarTitle').textContent = 'Edit Surat Keluar';
    document.getElementById('editKeluarId').value = record.id;
    document.getElementById('inNoAgendaKeluar').value = record.noAgenda ?? '';
    document.getElementById('inTglKirim').value = record.tglKirim || '';
    document.getElementById('inTglSuratKeluar').value = record.tglSurat || '';
    document.getElementById('inNoKeluar').value = record.noSurat || '';
    document.getElementById('inKepada').value = record.kepada || '';
    document.getElementById('inPerihalKeluar').value = record.perihal || '';
    document.getElementById('inIsiKeluar').value = record.isi || '';
    document.getElementById('inKodeArsip').value = record.kode || '';
    const file = document.getElementById('inFileKeluar');
    if (file) file.value = '';
    getModal('formSuratKeluarModal')?.show();
}

async function simpanSuratKeluar(event) {
    event?.preventDefault();
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat menyimpan surat keluar.');
    const id = document.getElementById('editKeluarId').value;
    const current = id ? suratKeluarData.find(item => String(item.id) === String(id)) : null;
    const file = document.getElementById('inFileKeluar')?.files?.[0] || null;
    let fileId = current?.fileId || null;
    let fileName = current?.fileName || '';
    if (file) {
        try {
            fileId = await saveUploadedFile(file);
            fileName = file.name;
        } catch (error) {
            console.error(error);
            return alert('File belum berhasil disimpan pada penyimpanan browser.');
        }
    }
    const values = {
        noAgenda: document.getElementById('inNoAgendaKeluar').value === '' ? (current?.noAgenda ?? nextAgendaNumber(suratKeluarData)) : Number(document.getElementById('inNoAgendaKeluar').value),
        tglKirim: document.getElementById('inTglKirim').value,
        tglSurat: document.getElementById('inTglSuratKeluar').value,
        noSurat: document.getElementById('inNoKeluar').value.trim(),
        kepada: document.getElementById('inKepada').value.trim(),
        perihal: document.getElementById('inPerihalKeluar').value.trim(),
        isi: document.getElementById('inIsiKeluar').value.trim(),
        kode: document.getElementById('inKodeArsip').value.trim(),
        fileId,
        fileName,
        source: current?.source || 'Input Manual'
    };
    if (current) Object.assign(current, values);
    else suratKeluarData.push({ id: nextRecordId(suratKeluarData), ...values });
    if (!saveDataToStorage()) return;
    renderAll();
    getModal('formSuratKeluarModal')?.hide();
    document.getElementById('formInputKeluar')?.reset();
    alert(current ? 'Surat keluar berhasil diperbarui.' : 'Surat keluar berhasil disimpan.');
}

async function hapusSuratMasuk(id) {
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat menghapus surat masuk.');
    const record = suratMasukData.find(item => String(item.id) === String(id));
    if (!record) return;
    if (!confirm('Hapus surat masuk ini beserta hubungan disposisinya?')) return;
    if (record.fileId) await deleteStoredFile(record.fileId).catch(console.error);
    suratMasukData = suratMasukData.filter(item => String(item.id) !== String(id));
    saveDataToStorage();
    renderAll();
}

async function hapusSuratKeluar(id) {
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat menghapus surat keluar.');
    const record = suratKeluarData.find(item => String(item.id) === String(id));
    if (!record) return;
    if (!confirm('Apakah Anda yakin ingin menghapus surat keluar ini?')) return;
    if (record.fileId) await deleteStoredFile(record.fileId).catch(console.error);
    suratKeluarData = suratKeluarData.filter(item => String(item.id) !== String(id));
    saveDataToStorage();
    renderAll();
}

/* ----------------------- Penyimpanan lampiran IndexedDB ----------------------- */
function openFileDatabase() {
    if (fileDbPromise) return fileDbPromise;
    fileDbPromise = new Promise((resolve, reject) => {
        if (!('indexedDB' in window)) return reject(new Error('IndexedDB tidak tersedia di browser ini.'));
        const request = indexedDB.open('HMAPD_Kearsipan_Files', 1);
        request.onupgradeneeded = () => {
            const db = request.result;
            if (!db.objectStoreNames.contains('lampiran')) db.createObjectStore('lampiran', { keyPath: 'id' });
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error || new Error('Database lampiran gagal dibuka.'));
    });
    return fileDbPromise;
}

const SUPABASE_STORAGE_BUCKET = 'lampiran-surat';

async function saveUploadedFile(file) {
    // Jika Supabase aktif, unggah ke Supabase Storage agar bisa dibuka dari HP & perangkat lain
    if (supabaseClient) {
        try {
            const safeBase = (file.name || 'lampiran').replace(/[^a-zA-Z0-9._-]/g, '_');
            const filePath = `${Date.now()}-${Math.random().toString(16).slice(2)}-${safeBase}`;
            const { error: uploadError } = await supabaseClient.storage
                .from(SUPABASE_STORAGE_BUCKET)
                .upload(filePath, file, { cacheControl: '3600', upsert: false });
            if (uploadError) throw uploadError;
            const { data: pubData } = supabaseClient.storage
                .from(SUPABASE_STORAGE_BUCKET)
                .getPublicUrl(filePath);
            if (pubData?.publicUrl) {
                return pubData.publicUrl;
            }
        } catch (err) {
            console.error('Gagal unggah ke Supabase Storage, beralih ke penyimpanan lokal:', err);
            alert('File gagal diunggah ke Supabase Storage (pastikan bucket "lampiran-surat" sudah dibuat & Public). Sementara file disimpan di browser lokal saja.');
        }
    }
    const db = await openFileDatabase();
    const id = `hmapd-file-${Date.now()}-${Math.random().toString(16).slice(2)}`;
    return new Promise((resolve, reject) => {
        const tx = db.transaction('lampiran', 'readwrite');
        tx.objectStore('lampiran').put({ id, name: file.name, type: file.type, size: file.size, blob: file, savedAt: new Date().toISOString() });
        tx.oncomplete = () => resolve(id);
        tx.onerror = () => reject(tx.error || new Error('Gagal menyimpan lampiran.'));
        tx.onabort = () => reject(tx.error || new Error('Penyimpanan lampiran dibatalkan.'));
    });
}

async function getStoredFile(id) {
    const db = await openFileDatabase();
    return new Promise((resolve, reject) => {
        const request = db.transaction('lampiran', 'readonly').objectStore('lampiran').get(id);
        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => reject(request.error || new Error('Lampiran tidak dapat dibaca.'));
    });
}

async function deleteStoredFile(id) {
    if (typeof id === 'string' && /^https?:\/\//i.test(id)) return true;
    const db = await openFileDatabase();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('lampiran', 'readwrite');
        tx.objectStore('lampiran').delete(id);
        tx.oncomplete = () => resolve(true);
        tx.onerror = () => reject(tx.error || new Error('Lampiran tidak dapat dihapus.'));
    });
}

function findFileRecord(type, id) {
    const records = type === 'masuk' ? suratMasukData : suratKeluarData;
    return records.find(item => String(item.id) === String(id));
}

async function previewFile(type, id) {
    const record = findFileRecord(type, id);
    const target = document.getElementById('previewContent');
    if (!record || !target) return;
    if (previewObjectUrl) URL.revokeObjectURL(previewObjectUrl);
    previewObjectUrl = null;
    if (!record.fileId) {
        target.innerHTML = '<div class="text-muted py-4"><i class="fa-solid fa-file-circle-xmark fa-3x mb-3"></i><p class="mb-0">Belum ada file lampiran pada surat ini.</p></div>';
        getModal('previewModal')?.show();
        return;
    }
    try {
        const safeName = escapeHTML(record.fileName || 'Lampiran');
        if (typeof record.fileId === 'string' && /^https?:\/\//i.test(record.fileId)) {
            const url = record.fileId;
            if (/\.(png|jpe?g|gif|webp|bmp)$/i.test(record.fileName || url)) {
                target.innerHTML = `<img src="${url}" alt="Preview ${safeName}" class="img-fluid rounded shadow-sm" style="max-height: 65vh;">`;
            } else if (/\.pdf$/i.test(record.fileName || url)) {
                target.innerHTML = `<iframe src="${url}" title="${safeName}" width="100%" height="600" class="border rounded"></iframe><div class="mt-2"><a href="${url}" target="_blank" rel="noopener" class="btn btn-outline-primary btn-sm">Buka PDF di Tab Baru</a></div>`;
            } else {
                target.innerHTML = `<p class="text-muted">Lampiran: <strong>${safeName}</strong></p><a href="${url}" target="_blank" rel="noopener" class="btn btn-primary btn-sm">Buka / Unduh File</a>`;
            }
            getModal('previewModal')?.show();
            return;
        }
        const file = await getStoredFile(record.fileId);
        if (!file?.blob) throw new Error('Data file tidak ditemukan di IndexedDB.');
        previewObjectUrl = URL.createObjectURL(file.blob);
        const localSafeName = escapeHTML(file.name || record.fileName || 'Lampiran');
        if ((file.type || '').startsWith('image/')) {
            target.innerHTML = `<img src="${previewObjectUrl}" alt="Preview ${localSafeName}" class="img-fluid rounded shadow-sm" style="max-height: 65vh;">`;
        } else if (file.type === 'application/pdf' || /\.pdf$/i.test(file.name || '')) {
            target.innerHTML = `<iframe src="${previewObjectUrl}" title="${localSafeName}" width="100%" height="600" class="border rounded"></iframe>`;
        } else {
            target.innerHTML = `<p class="text-muted">Lampiran: <strong>${localSafeName}</strong></p><button class="btn btn-primary btn-sm" type="button" onclick="downloadFile('${type}', ${Number(id)})">Unduh file</button>`;
        }
        getModal('previewModal')?.show();
    } catch (error) {
        console.error(error);
        target.innerHTML = '<p class="text-danger">File tidak ditemukan atau tidak dapat dibuka. Unggah ulang file tersebut.</p>';
        getModal('previewModal')?.show();
    }
}

async function downloadFile(type, id) {
    const record = findFileRecord(type, id);
    if (!record?.fileId) return alert('Belum ada file lampiran untuk surat ini.');
    if (typeof record.fileId === 'string' && /^https?:\/\//i.test(record.fileId)) {
        const anchor = document.createElement('a');
        anchor.href = record.fileId;
        anchor.target = '_blank';
        anchor.rel = 'noopener';
        anchor.download = record.fileName || 'lampiran';
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        return;
    }
    try {
        const file = await getStoredFile(record.fileId);
        if (!file?.blob) throw new Error('File tidak ditemukan.');
        const url = URL.createObjectURL(file.blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = file.name || record.fileName || 'lampiran';
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        window.setTimeout(() => URL.revokeObjectURL(url), 1500);
    } catch (error) {
        console.error(error);
        alert('File tidak dapat diunduh. Silakan unggah ulang lampiran.');
    }
}

/* ----------------------- Alur lembar disposisi ----------------------- */
function statusDisposisiSurat(surat) {
    if (!surat?.disposisi) return 'Belum Dibuat';
    return surat.disposisi.status || 'Dalam Proses';
}

function setText(id, value) {
    const element = document.getElementById(id);
    if (element) element.textContent = value || '-';
}

function normaliseSifatCode(val) {
    const raw = String(val ?? '').trim().toUpperCase();
    if (raw === 'P' || raw === 'PENTING') return 'P';
    if (raw === 'B' || raw === 'BIASA') return 'B';
    if (raw === 'R' || raw === 'RAHASIA') return 'R';
    return '';
}

function syncInstruksiDropdown() {
    const sel = document.getElementById('dspInstruksiSelect');
    const txt = document.getElementById('dspInstruksi');
    if (sel && txt && sel.value) {
        txt.value = sel.value;
    }
}

function updateDisposisiSheet(surat) {
    const data = surat.disposisi || {};
    setText('dspPerihal', surat.perihal);
    setText('dspTglSelesaiText', formatTanggalDisposisi(data.tanggalSelesai));
    setText('dspTglSurat', formatTanggalDisposisi(surat.tglSurat || surat.tglTerima));
    setText('dspNoSurat', surat.noSurat);
    setText('dspPengirim', surat.pengirim);
    const lampiranVal = surat.lampiran || data.lampiran || '-';
    setText('dspLampiran', lampiranVal);
    setText('dspInstruksiText', data.tindakan || '-');

    // Lingkari P / B / R pada pojok kanan atas sesuai sifat surat dari surat masuk
    const sifatCode = normaliseSifatCode(surat.sifatSurat || data.prioritas);
    ['P', 'B', 'R'].forEach(code => {
        const el = document.getElementById(`pbrMark${code}`);
        if (el) el.classList.toggle('circled', sifatCode === code);
    });

    // Daftar 1..4 Diteruskan dengan garis titik-titik untuk baris yang kosong
    const receiverList = document.getElementById('dspPenerimaList');
    if (receiverList) {
        const recipients = Array.isArray(data.penerima) ? data.penerima : (data.penerima ? [data.penerima] : []);
        const dots = '........................................................................................';
        let itemsHtml = '';
        for (let i = 0; i < 4; i++) {
            const val = recipients[i] ? escapeHTML(recipients[i]) : dots;
            itemsHtml += `<li>${val}</li>`;
        }
        receiverList.innerHTML = itemsHtml;
    }
}

function bukaDisposisi(id) {
    selectedSuratId = Number(id);
    const surat = suratMasukData.find(item => String(item.id) === String(id));
    if (!surat) return;
    const hasDisposisi = Boolean(surat.disposisi);
    const data = surat.disposisi || {};

    // Isi form khusus Admin (identitas surat & sifat surat)
    const admPerihal = document.getElementById('admDspPerihalInput');
    if (admPerihal) admPerihal.value = surat.perihal || '';
    const admSifat = document.getElementById('admDspSifatInput');
    if (admSifat) admSifat.value = normaliseSifatCode(surat.sifatSurat || data.prioritas);
    const admTglSurat = document.getElementById('admDspTglSuratInput');
    if (admTglSurat) admTglSurat.value = surat.tglSurat || surat.tglTerima || '';
    const admNoSurat = document.getElementById('admDspNoSuratInput');
    if (admNoSurat) admNoSurat.value = surat.noSurat || '';
    const admPengirim = document.getElementById('admDspPengirimInput');
    if (admPengirim) admPengirim.value = surat.pengirim || '';
    const admLampiran = document.getElementById('admDspLampiranInput');
    if (admLampiran) admLampiran.value = surat.lampiran || data.lampiran || '';

    // Isi form khusus Pimpinan (tanggal penyelesaian, diteruskan kepada, instruksi)
    const tglSelesaiEl = document.getElementById('dspTglSelesai');
    if (tglSelesaiEl) tglSelesaiEl.value = data.tanggalSelesai || '';
    const statusEl = document.getElementById('dspStatusDisposisi');
    if (statusEl) statusEl.value = data.status || 'Dalam Proses';
    const instruksiEl = document.getElementById('dspInstruksi');
    if (instruksiEl) instruksiEl.value = data.tindakan || '';
    const instruksiSel = document.getElementById('dspInstruksiSelect');
    if (instruksiSel) {
        const matchOpt = Array.from(instruksiSel.options).some(opt => opt.value === (data.tindakan || ''));
        instruksiSel.value = matchOpt ? (data.tindakan || '') : '';
    }

    populateDisposisiPenerimaOptions();
    const recipientSelect = document.getElementById('dspPenerima');
    const savedRecipients = Array.isArray(data.penerima) ? data.penerima : (data.penerima ? [data.penerima] : []);
    if (recipientSelect) {
        Array.from(recipientSelect.options).forEach(option => {
            option.selected = savedRecipients.includes(option.value);
        });
    }

    updateDisposisiSheet(surat);
    applyRolePermissions();
    const saveButton = document.getElementById('btnSimpanDisposisi');
    if (saveButton) saveButton.textContent = hasDisposisi ? 'Simpan Perubahan Disposisi' : 'Simpan Disposisi';
    getModal('disposisiModal')?.show();
}

function simpanIdentitasDisposisiAdmin() {
    if (currentRole !== 'admin') return alert('Hanya Admin yang boleh mengubah identitas surat.');
    const surat = suratMasukData.find(item => String(item.id) === String(selectedSuratId));
    if (!surat) return alert('Surat masuk tidak ditemukan.');

    surat.perihal = document.getElementById('admDspPerihalInput').value.trim();
    surat.sifatSurat = document.getElementById('admDspSifatInput').value;
    surat.tglSurat = document.getElementById('admDspTglSuratInput').value;
    surat.noSurat = document.getElementById('admDspNoSuratInput').value.trim();
    surat.pengirim = document.getElementById('admDspPengirimInput').value.trim();
    surat.lampiran = document.getElementById('admDspLampiranInput').value.trim();

    if (!saveDataToStorage()) return;
    updateDisposisiSheet(surat);
    renderAll();
    alert('Identitas surat dan sifat surat berhasil diperbarui oleh Admin.');
}

function renderDaftarDisposisi() {
    const tbody = document.getElementById('tbodyDaftarDisposisi');
    if (!tbody) return;
    const sourceRows = Array.isArray(suratMasukData) ? suratMasukData : [];
    const keyword = normalizeSearchText(document.getElementById('searchDaftarDisposisi')?.value || '');
    const rows = sourceRows.filter(surat => normalizeSearchText(collectSearchableText({
        ...surat,
        statusDisposisi: statusDisposisiSurat(surat),
        penerimaDisposisi: surat.disposisi?.penerima,
        tindakanDisposisi: surat.disposisi?.tindakan
    })).includes(keyword));
    tbody.innerHTML = rows.length ? rows.map((surat, index) => {
        const status = statusDisposisiSurat(surat);
        const badge = status === 'Selesai' ? 'bg-success-subtle text-success' : status === 'Dalam Proses' ? 'bg-warning-subtle text-warning' : 'bg-secondary-subtle text-secondary';
        const action = currentRole === 'pimpinan' ? (surat.disposisi ? 'Edit Disposisi' : 'Buat Disposisi') : (currentRole === 'admin' ? 'Lihat / Edit Identitas' : 'Lihat');
        return `<tr>
            <td>${index + 1}</td>
            <td class="fw-semibold">${textOrDash(surat.noSurat)}</td>
            <td>${textOrDash(surat.pengirim)}</td>
            <td>${textOrDash(surat.perihal)}</td>
            <td>${formatTanggal(surat.tglTerima)}</td>
            <td><span class="badge ${badge}">${escapeHTML(status)}</span></td>
            <td><button type="button" class="btn btn-sm btn-outline-primary" onclick="bukaDisposisi(${Number(surat.id)})">${action}</button></td>
        </tr>`;
    }).join('') : '<tr><td colspan="7" class="text-center text-muted py-4">Belum ada surat yang sesuai.</td></tr>';
}

function filterDaftarDisposisi() {
    renderDaftarDisposisi();
}

function simpanDisposisiData() {
    if (currentRole !== 'pimpinan') return alert('Hanya Pimpinan yang boleh mengisi atau mengubah instruksi, tanggal penyelesaian, dan tujuan disposisi.');
    const surat = suratMasukData.find(item => String(item.id) === String(selectedSuratId));
    if (!surat) return alert('Surat masuk yang dipilih tidak ditemukan.');
    const penerima = Array.from(document.getElementById('dspPenerima').selectedOptions).map(option => option.value);
    const tindakan = document.getElementById('dspInstruksi').value.trim();
    if (!penerima.length || !tindakan) return alert('Pilih minimal satu tujuan (Diteruskan Kepada) dan isi Instruksi Pimpinan terlebih dahulu.');
    if (penerima.length > 4) return alert('Maksimal 4 bagian/manajer pada daftar Diteruskan Kepada.');

    surat.disposisi = {
        tanggalSelesai: document.getElementById('dspTglSelesai').value,
        penerima,
        tindakan,
        prioritas: surat.sifatSurat || surat.disposisi?.prioritas || '',
        lampiran: surat.lampiran || surat.disposisi?.lampiran || '-',
        status: document.getElementById('dspStatusDisposisi').value || 'Dalam Proses',
        diperbaruiPada: new Date().toISOString()
    };
    if (!saveDataToStorage()) return;
    updateDisposisiSheet(surat);
    renderDaftarDisposisi();
    alert('Lembar disposisi berhasil disimpan.');
}

function printDisposisiSheet() {
    if (currentRole === 'tamu') return alert('Akses cetak disposisi tidak tersedia untuk Tamu.');
    window.print();
}

async function downloadDisposisiPDF() {
    if (currentRole === 'tamu') return alert('Akses unduh disposisi tidak tersedia untuk Tamu.');
    const surat = suratMasukData.find(item => String(item.id) === String(selectedSuratId));
    if (!surat) return;
    const sheet = document.getElementById('disposisiSheet');
    if (!sheet) return;
    const safeNumber = String(surat.noSurat || `agenda-${surat.noAgenda || surat.id}`).replace(/[^a-z0-9_-]/gi, '_');
    const filename = `Lembar_Disposisi_${safeNumber}.pdf`;
    if (typeof window.html2pdf === 'function') {
        const options = {
            margin: 0.35,
            filename,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2, useCORS: true, backgroundColor: '#ffffff' },
            jsPDF: { unit: 'in', format: 'a4', orientation: 'portrait' }
        };
        await window.html2pdf().set(options).from(sheet).save();
    } else {
        alert('Pustaka PDF belum termuat. Gunakan fitur cetak browser sebagai alternatif.');
        window.print();
    }
}

/* ----------------------- Laporan Buku Agenda Surat Masuk / Keluar ----------------------- */
let activeLaporanType = 'masuk';

function bukaLaporanAgenda(type) {
    activeLaporanType = type;
    const titleEl = document.getElementById('laporanAgendaModalTitle');
    const headerEl = document.getElementById('laporanAgendaHeader');
    const subEl = document.getElementById('laporanAgendaSubtitle');
    const thead = document.getElementById('laporanAgendaThead');
    const tbody = document.getElementById('laporanAgendaTbody');

    const nowText = new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' });
    if (subEl) subEl.textContent = `Dicetak pada: ${nowText}`;

    if (type === 'masuk') {
        if (titleEl) titleEl.innerHTML = '<i class="fa-solid fa-book me-2"></i>Laporan Buku Agenda Surat Masuk';
        if (headerEl) headerEl.textContent = 'LAPORAN BUKU AGENDA SURAT MASUK';
        if (thead) {
            thead.innerHTML = `<tr>
                <th>No Agenda</th>
                <th>Tgl Terima</th>
                <th>No Surat</th>
                <th>Tgl Surat</th>
                <th>Dari</th>
                <th>Perihal</th>
                <th>Isi Ringkas</th>
                <th>Kode Arsip</th>
            </tr>`;
        }
        if (tbody) {
            tbody.innerHTML = suratMasukData.length ? suratMasukData.map(item => `<tr>
                <td>${textOrDash(item.noAgenda)}</td>
                <td>${formatTanggal(item.tglTerima)}</td>
                <td>${textOrDash(item.noSurat)}</td>
                <td>${formatTanggal(item.tglSurat)}</td>
                <td>${textOrDash(item.pengirim)}</td>
                <td>${textOrDash(item.perihal)}</td>
                <td>${textOrDash(item.isi)}</td>
                <td>${textOrDash(item.kodeArsip)}</td>
            </tr>`).join('') : '<tr><td colspan="8" class="text-center">Tidak ada data surat masuk.</td></tr>';
        }
    } else {
        if (titleEl) titleEl.innerHTML = '<i class="fa-solid fa-book me-2"></i>Laporan Buku Agenda Surat Keluar';
        if (headerEl) headerEl.textContent = 'LAPORAN BUKU AGENDA SURAT KELUAR';
        if (thead) {
            thead.innerHTML = `<tr>
                <th>No Agenda</th>
                <th>Tgl Kirim</th>
                <th>No Surat</th>
                <th>Tgl Surat</th>
                <th>Kepada</th>
                <th>Perihal</th>
                <th>Isi Ringkas</th>
                <th>Kode Arsip</th>
            </tr>`;
        }
        if (tbody) {
            tbody.innerHTML = suratKeluarData.length ? suratKeluarData.map(item => `<tr>
                <td>${textOrDash(item.noAgenda)}</td>
                <td>${formatTanggal(item.tglKirim)}</td>
                <td>${textOrDash(item.noSurat)}</td>
                <td>${formatTanggal(item.tglSurat)}</td>
                <td>${textOrDash(item.kepada)}</td>
                <td>${textOrDash(item.perihal)}</td>
                <td>${textOrDash(item.isi)}</td>
                <td>${textOrDash(item.kode)}</td>
            </tr>`).join('') : '<tr><td colspan="8" class="text-center">Tidak ada data surat keluar.</td></tr>';
        }
    }

    getModal('laporanAgendaModal')?.show();
}

function printLaporanAgenda() {
    const area = document.getElementById('laporanAgendaPrintable');
    if (!area) return;
    const win = window.open('', '_blank', 'width=1000,height=700');
    if (!win) return window.print();
    win.document.write(`<!DOCTYPE html><html><head><title>Laporan Buku Agenda PT HMAPD</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>body{padding:24px;font-family:'Times New Roman',serif;} table{width:100%;border-collapse:collapse;} th,td{border:1px solid #000 !important;padding:6px 8px;font-size:12px;}</style>
        </head><body>${area.innerHTML}</body></html>`);
    win.document.close();
    win.focus();
    setTimeout(() => { win.print(); win.close(); }, 300);
}

async function downloadLaporanAgendaPDF() {
    const area = document.getElementById('laporanAgendaPrintable');
    if (!area) return;
    const filename = activeLaporanType === 'masuk' ? 'Laporan_Buku_Agenda_Surat_Masuk.pdf' : 'Laporan_Buku_Agenda_Surat_Keluar.pdf';
    if (typeof window.html2pdf === 'function') {
        const options = {
            margin: 0.3,
            filename,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2, useCORS: true, backgroundColor: '#ffffff' },
            jsPDF: { unit: 'in', format: 'a4', orientation: 'landscape' }
        };
        await window.html2pdf().set(options).from(area).save();
    } else {
        printLaporanAgenda();
    }
}


/* ----------------------- Pengaturan Kontak Admin & WhatsApp 1-Klik ----------------------- */
function normaliseWaNumber(num) {
    let clean = String(num ?? '').replace(/[^0-9]/g, '');
    if (clean.startsWith('0')) clean = '62' + clean.slice(1);
    return clean;
}

function populateDisposisiPenerimaOptions() {
    const select = document.getElementById('dspPenerima');
    if (!select) return;
    const prevSelected = Array.from(select.selectedOptions).map(o => o.value);
    select.innerHTML = kontakManajerData.map(item =>
        `<option value="${escapeHTML(item.nama)}">${escapeHTML(item.nama)}</option>`
    ).join('');
    Array.from(select.options).forEach(opt => {
        opt.selected = prevSelected.includes(opt.value);
    });
}

function renderKontakSettings() {
    const namaEl = document.getElementById('cfgAdminNama');
    const waEl = document.getElementById('cfgAdminWa');
    if (namaEl) namaEl.value = kontakAdminData.nama || '';
    if (waEl) waEl.value = kontakAdminData.wa || '';

    const tbody = document.getElementById('tbodyKontakManajer');
    if (!tbody) return;
    tbody.innerHTML = kontakManajerData.length ? kontakManajerData.map((item, idx) => `
        <tr>
            <td>${idx + 1}</td>
            <td class="fw-semibold">${escapeHTML(item.nama)}</td>
            <td>${item.wa ? escapeHTML(item.wa) : '<span class="text-muted small">Belum diisi</span>'}</td>
            <td>
                <button type="button" class="btn btn-sm btn-outline-warning mb-1" onclick="bukaModalEditKontak(${Number(item.id)})"><i class="fa-solid fa-pen-to-square"></i> Edit</button>
                <button type="button" class="btn btn-sm btn-outline-danger mb-1" onclick="hapusKontakManajer(${Number(item.id)})"><i class="fa-solid fa-trash"></i> Hapus</button>
            </td>
        </tr>
    `).join('') : '<tr><td colspan="4" class="text-center text-muted py-3">Belum ada daftar manajer.</td></tr>';
}

function simpanKontakAdmin(event) {
    event?.preventDefault();
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat mengubah kontak.');
    kontakAdminData = {
        nama: document.getElementById('cfgAdminNama').value.trim(),
        wa: normaliseWaNumber(document.getElementById('cfgAdminWa').value)
    };
    saveDataToStorage();
    renderKontakSettings();
    alert('Kontak WhatsApp Admin berhasil disimpan.');
    return false;
}

function bukaModalTambahKontak() {
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat menambah tujuan disposisi.');
    document.getElementById('modalKontakTitle').textContent = 'Tambah Tujuan Disposisi';
    document.getElementById('editKontakId').value = '';
    document.getElementById('formKontakManajer')?.reset();
    getModal('formKontakModal')?.show();
}

function bukaModalEditKontak(id) {
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat mengedit tujuan disposisi.');
    const item = kontakManajerData.find(x => String(x.id) === String(id));
    if (!item) return;
    document.getElementById('modalKontakTitle').textContent = 'Edit Tujuan Disposisi';
    document.getElementById('editKontakId').value = item.id;
    document.getElementById('inNamaManajer').value = item.nama || '';
    document.getElementById('inWaManajer').value = item.wa || '';
    getModal('formKontakModal')?.show();
}

function simpanKontakManajer(event) {
    event?.preventDefault();
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat menyimpan kontak manajer.');
    const id = document.getElementById('editKontakId').value;
    const nama = document.getElementById('inNamaManajer').value.trim();
    const wa = normaliseWaNumber(document.getElementById('inWaManajer').value);
    if (!nama) return alert('Nama bagian / manajer wajib diisi.');

    if (id) {
        const existing = kontakManajerData.find(x => String(x.id) === String(id));
        if (existing) {
            existing.nama = nama;
            existing.wa = wa;
        }
    } else {
        kontakManajerData.push({ id: nextRecordId(kontakManajerData), nama, wa });
    }
    saveDataToStorage();
    renderKontakSettings();
    populateDisposisiPenerimaOptions();
    getModal('formKontakModal')?.hide();
    return false;
}

function hapusKontakManajer(id) {
    if (currentRole !== 'admin') return alert('Hanya Admin yang dapat menghapus kontak.');
    if (!confirm('Hapus tujuan disposisi ini?')) return;
    kontakManajerData = kontakManajerData.filter(x => String(x.id) !== String(id));
    saveDataToStorage();
    renderKontakSettings();
    populateDisposisiPenerimaOptions();
}

function kirimWhatsAppDisposisi() {
    if (currentRole !== 'pimpinan') return alert('Hanya Pimpinan yang dapat mengirim notifikasi WhatsApp disposisi.');
    const surat = suratMasukData.find(item => String(item.id) === String(selectedSuratId));
    if (!surat?.disposisi) return alert('Simpan disposisi terlebih dahulu sebelum mengirim WhatsApp.');

    const data = surat.disposisi;
    const recipients = Array.isArray(data.penerima) ? data.penerima : (data.penerima ? [data.penerima] : []);
    const tglSelesai = formatTanggalDisposisi(data.tanggalSelesai);
    const instruksi = data.tindakan || '-';
    const targets = [];

    // 1. Target Admin
    const adminWa = normaliseWaNumber(kontakAdminData.wa);
    const adminMsg = `Halo ${kontakAdminData.nama || 'Admin'}, lembar disposisi untuk surat dari *${surat.pengirim || '-'}* (No: ${surat.noSurat || '-'}) perihal *${surat.perihal || '-'}* telah selesai diproses oleh Pimpinan.\n\nDiteruskan kepada: ${recipients.join(', ') || '-'}\nInstruksi: ${instruksi}\nTanggal Penyelesaian: ${tglSelesai}\nStatus: ${data.status || 'Dalam Proses'}`;
    targets.push({
        label: `Admin (${kontakAdminData.nama || 'Admin'})`,
        wa: adminWa,
        url: adminWa ? `https://wa.me/${adminWa}?text=${encodeURIComponent(adminMsg)}` : ''
    });

    // 2. Target Manajer/Bagian terpilih
    recipients.forEach(namaBagian => {
        const kontak = kontakManajerData.find(k => normaliseIdentity(k.nama) === normaliseIdentity(namaBagian));
        const waNum = normaliseWaNumber(kontak?.wa || '');
        const msg = `Yth. *${namaBagian}*,\nAnda menerima disposisi baru dari Pimpinan PT HMAPD:\n\nAsal Surat: ${surat.pengirim || '-'}\nNomor Surat: ${surat.noSurat || '-'}\nPerihal: *${surat.perihal || '-'}*\nInstruksi Pimpinan: *${instruksi}*\nTanggal Penyelesaian: ${tglSelesai}\n\nMohon segera ditindaklanjuti. Terima kasih.`;
        targets.push({
            label: namaBagian,
            wa: waNum,
            url: waNum ? `https://wa.me/${waNum}?text=${encodeURIComponent(msg)}` : ''
        });
    });

    // Coba buka otomatis semua penerima yang memiliki nomor WA
    targets.forEach(t => {
        if (t.url) window.open(t.url, '_blank');
    });

    // Tampilkan daftar tombol antrian di modal sebagai cadangan jika browser memblokir popup ganda
    const listEl = document.getElementById('waDispatchList');
    if (listEl) {
        listEl.innerHTML = targets.map(t => t.url
            ? `<a href="${t.url}" target="_blank" rel="noopener" class="btn btn-success d-flex justify-content-between align-items-center"><span><i class="fa-brands fa-whatsapp me-2"></i>Kirim ke ${escapeHTML(t.label)}</span><span class="small opacity-75">${escapeHTML(t.wa)}</span></a>`
            : `<div class="alert alert-warning py-2 px-3 mb-0 small"><i class="fa-solid fa-triangle-exclamation me-1"></i>Nomor WA untuk <strong>${escapeHTML(t.label)}</strong> belum diisi oleh Admin di menu Pengaturan Kontak.</div>`
        ).join('');
    }
    getModal('waDispatchModal')?.show();
}

/* ----------------------- Inisialisasi halaman ----------------------- */
document.addEventListener('DOMContentLoaded', () => {
    togglePasswordInput();
    renderAll();
    if (initSupabase()) {
        syncFromSupabase(false);
    }
    try {
        if (localStorage.getItem(STORAGE_KEYS.tema) === 'dark') {
            document.body.classList.add('dark-mode');
            const icon = document.getElementById('themeIcon');
            if (icon) icon.className = 'fa-solid fa-sun';
        }
    } catch (_) {}

    document.getElementById('previewModal')?.addEventListener('hidden.bs.modal', () => {
        if (previewObjectUrl) URL.revokeObjectURL(previewObjectUrl);
        previewObjectUrl = null;
        const target = document.getElementById('previewContent');
        if (target) target.innerHTML = '';
    });
});
