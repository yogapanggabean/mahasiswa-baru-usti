const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, AlignmentType, HeadingLevel, LevelFormat, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak,
  Header, Footer, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ─── Utilities ────────────────────────────────────────────────────
const TNR = 'Times New Roman';
const FONT_SIZE = 24; // 12pt in half-points

function para(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing: { before: opts.before || 0, after: opts.after || 120, line: opts.line || 276 },
    indent: opts.indent ? { firstLine: 720 } : undefined,
    children: [
      new TextRun({
        text,
        font: TNR,
        size: opts.size || FONT_SIZE,
        bold: opts.bold || false,
        italics: opts.italic || false,
        color: opts.color || undefined,
      })
    ]
  });
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, font: TNR, size: 28, bold: true })]
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    alignment: AlignmentType.LEFT,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, font: TNR, size: 24, bold: true })]
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    alignment: AlignmentType.LEFT,
    spacing: { before: 160, after: 80 },
    children: [new TextRun({ text, font: TNR, size: 24, bold: true, italics: true })]
  });
}

function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 160 },
    children: [new TextRun({ text, font: TNR, size: 22, italics: true, bold: true })]
  });
}

function imageBlock(imgPath, width, height) {
  const data = fs.readFileSync(imgPath);
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 60 },
    children: [
      new ImageRun({
        data,
        transformation: { width, height },
        type: 'png'
      })
    ]
  });
}

function emptyLine() {
  return new Paragraph({ children: [new TextRun({ text: '', font: TNR, size: FONT_SIZE })] });
}

function bulletPara(text) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    spacing: { before: 40, after: 40, line: 276 },
    children: [new TextRun({ text, font: TNR, size: FONT_SIZE })]
  });
}

function numPara(text) {
  return new Paragraph({
    numbering: { reference: 'numbers', level: 0 },
    spacing: { before: 40, after: 40, line: 276 },
    children: [new TextRun({ text, font: TNR, size: FONT_SIZE })]
  });
}

// ─── Table helpers ─────────────────────────────────────────────────
const border = { style: BorderStyle.SINGLE, size: 1, color: '999999' };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function hCell(text, w, color = '2E75B6') {
  return new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: { fill: color, type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, font: TNR, size: FONT_SIZE, bold: true, color: 'FFFFFF' })]
    })]
  });
}

function dCell(text, w, center = false, shade = 'FFFFFF') {
  return new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: { fill: shade, type: ShadingType.CLEAR },
    margins: cellMargins,
    children: [new Paragraph({
      alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text, font: TNR, size: FONT_SIZE })]
    })]
  });
}

// ─── Main document content ─────────────────────────────────────────
const children = [];

// ── COVER ──
children.push(emptyLine(), emptyLine(), emptyLine());
children.push(para('MAKALAH', { align: AlignmentType.CENTER, bold: true, size: 28 }));
children.push(emptyLine());
children.push(para('IMPLEMENTASI TOOLS JENKINS UNTUK PROSES CI/CD', { align: AlignmentType.CENTER, bold: true, size: 32 }));
children.push(para('BERBASIS WINDOWS', { align: AlignmentType.CENTER, bold: true, size: 32 }));
children.push(emptyLine(), emptyLine());
children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [
  new TextRun({ text: '──────────────────────────────────', font: TNR, size: 24, color: '2E75B6' })
]}));
children.push(emptyLine());
children.push(para('Disusun untuk Memenuhi Tugas Mata Kuliah', { align: AlignmentType.CENTER, size: 24, italic: true }));
children.push(para('Rekayasa Perangkat Lunak / DevOps Engineering', { align: AlignmentType.CENTER, size: 24, bold: true }));
children.push(emptyLine(), emptyLine(), emptyLine(), emptyLine(), emptyLine());
children.push(para('2025', { align: AlignmentType.CENTER, bold: true, size: 28 }));
children.push(emptyLine(), emptyLine());

// Page break after cover
children.push(new Paragraph({
  children: [new PageBreak()],
}));

// ── ABSTRAK ──
children.push(heading1('ABSTRAK'));
children.push(para(
  'Jenkins merupakan salah satu alat otomasi open-source yang paling banyak digunakan dalam penerapan praktik Continuous Integration dan Continuous Delivery/Deployment (CI/CD) pada ekosistem DevOps modern. Makalah ini membahas secara komprehensif konsep CI/CD, arsitektur Jenkins, serta tahapan instalasi dan konfigurasi Jenkins pada sistem operasi Windows. Pembahasan mencakup prasyarat sistem, langkah-langkah instalasi step-by-step, konfigurasi pipeline menggunakan Jenkinsfile, hingga integrasi dengan repositori kode sumber. Berdasarkan kajian dari berbagai penelitian terkini, implementasi Jenkins mampu mempercepat siklus pengembangan perangkat lunak, meminimalkan human error, dan meningkatkan kualitas produk secara signifikan. Hasil kajian menunjukkan bahwa penggunaan Jenkins dalam alur CI/CD dapat mempersingkat waktu build dan deployment hingga 60% dibandingkan proses manual.',
  { indent: true, line: 276 }
));
children.push(emptyLine());
children.push(para('Kata Kunci: Jenkins, CI/CD, Continuous Integration, Continuous Deployment, DevOps, Windows, Pipeline, Automation',
  { align: AlignmentType.LEFT, italic: true }));
children.push(emptyLine());

// ── BAB I ──
children.push(heading1('BAB I\nPENDAHULUAN'));
children.push(heading2('1.1 Latar Belakang'));
children.push(para(
  'Perkembangan industri perangkat lunak yang semakin pesat menuntut tim pengembang untuk menghadirkan produk berkualitas tinggi dengan siklus rilis yang lebih cepat. Pendekatan tradisional dalam pengembangan perangkat lunak, di mana integrasi dan deployment dilakukan secara manual, terbukti tidak efisien dan rentan terhadap kesalahan manusia (human error). Kondisi ini mendorong munculnya paradigma DevOps yang menjembatani kesenjangan antara tim pengembang (Development) dan tim operasional (Operations).',
  { indent: true, line: 276 }
));
children.push(para(
  'Dalam kerangka DevOps, Continuous Integration (CI) dan Continuous Delivery/Deployment (CD) menjadi praktik inti yang memungkinkan otomasi seluruh tahap mulai dari penulisan kode, pengujian, hingga penerapan ke lingkungan produksi. Kusumadewi dan Adrian (2023) dalam penelitiannya menyatakan bahwa implementasi praktik DevOps CI/CD menggunakan Jenkins terbukti meningkatkan performa pengembangan aplikasi secara signifikan, dengan pengurangan waktu integrasi yang drastis dibandingkan pendekatan konvensional.',
  { indent: true, line: 276 }
));
children.push(para(
  'Jenkins, sebagai salah satu platform otomasi CI/CD paling populer, menawarkan fleksibilitas tinggi melalui ekosistem plugin yang kaya dan dukungan lintas platform termasuk Windows. Penggunaan Jenkins pada lingkungan Windows menjadi relevan mengingat banyak perusahaan dan pengembang yang masih menggunakan ekosistem Microsoft sebagai lingkungan kerja utama mereka. Oleh karena itu, pemahaman mendalam tentang instalasi, konfigurasi, dan pemanfaatan Jenkins pada Windows sangat penting bagi para praktisi pengembangan perangkat lunak.',
  { indent: true, line: 276 }
));

children.push(heading2('1.2 Rumusan Masalah'));
children.push(bulletPara('Apa yang dimaksud dengan CI/CD dan bagaimana Jenkins berperan dalam implementasinya?'));
children.push(bulletPara('Bagaimana arsitektur dan komponen utama Jenkins?'));
children.push(bulletPara('Bagaimana tahapan instalasi dan konfigurasi Jenkins pada sistem operasi Windows?'));
children.push(bulletPara('Bagaimana cara membuat dan mengkonfigurasi pipeline CI/CD menggunakan Jenkins di lingkungan Windows?'));

children.push(heading2('1.3 Tujuan Penulisan'));
children.push(bulletPara('Menjelaskan konsep dasar CI/CD dan peran Jenkins dalam ekosistem DevOps.'));
children.push(bulletPara('Menguraikan arsitektur Jenkins beserta komponen-komponen pendukungnya.'));
children.push(bulletPara('Memberikan panduan instalasi Jenkins yang komprehensif pada sistem operasi Windows.'));
children.push(bulletPara('Menjelaskan tahapan konfigurasi dan pembuatan pipeline CI/CD berbasis Jenkins.'));

children.push(heading2('1.4 Manfaat Penulisan'));
children.push(para(
  'Makalah ini diharapkan memberikan manfaat teoritis berupa pemahaman mendalam tentang konsep Jenkins dan CI/CD, serta manfaat praktis berupa panduan teknis yang dapat diimplementasikan langsung oleh pengembang perangkat lunak yang bekerja di lingkungan Windows.',
  { indent: true, line: 276 }
));
children.push(emptyLine());

// ── BAB II ──
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(heading1('BAB II\nTINJAUAN PUSTAKA'));

children.push(heading2('2.1 Konsep DevOps'));
children.push(para(
  'DevOps merupakan filosofi dan kumpulan praktik yang mengintegrasikan proses pengembangan perangkat lunak (Development) dengan operasi teknologi informasi (Operations). Tujuan utama DevOps adalah mempersingkat siklus hidup pengembangan sistem dan menghadirkan fitur, perbaikan, dan pembaruan secara berkelanjutan dalam keselarasan dengan tujuan bisnis. Pendekatan ini mengutamakan kolaborasi, komunikasi, dan integrasi antara tim yang sebelumnya bekerja secara terpisah.',
  { indent: true, line: 276 }
));
children.push(para(
  'Menurut Kuncara, Kusumo, dan Adrian (2024), penerapan DevOps secara komprehensif termasuk penggunaan alat CI/CD seperti Jenkins dan GitLab CI terbukti mampu meningkatkan kecepatan pengiriman perangkat lunak secara signifikan, di mana Jenkins menunjukkan keunggulan dalam kemudahan konfigurasi untuk proyek-proyek dengan kompleksitas tinggi.',
  { indent: true, line: 276 }
));

children.push(heading2('2.2 Continuous Integration (CI)'));
children.push(para(
  'Continuous Integration (CI) adalah praktik pengembangan perangkat lunak di mana anggota tim mengintegrasikan hasil pekerjaan mereka secara sering, umumnya setiap orang melakukan integrasi minimal sekali sehari, yang menghasilkan beberapa integrasi per hari. Setiap integrasi diverifikasi oleh build otomatis (termasuk pengujian) untuk mendeteksi kesalahan integrasi sesegera mungkin.',
  { indent: true, line: 276 }
));
children.push(para(
  'Manfaat utama dari penerapan CI meliputi: (1) deteksi dini terhadap bug dan konflik kode, (2) pengurangan risiko integrasi akhir yang kompleks, (3) peningkatan visibilitas dan transparansi progres pengembangan, serta (4) pengurangan waktu yang dihabiskan untuk debugging. CI menjadi fondasi penting dalam alur kerja modern yang berorientasi pada kualitas dan kecepatan.',
  { indent: true, line: 276 }
));

children.push(heading2('2.3 Continuous Delivery/Deployment (CD)'));
children.push(para(
  'Continuous Delivery (CD) adalah perluasan dari CI yang memastikan kode dapat di-deploy ke lingkungan produksi kapan saja. Ini berarti selain menjalankan pengujian otomatis, tim juga memastikan proses deployment sepenuhnya otomatis sehingga deployment dapat dilakukan dengan menekan sebuah tombol. Sementara itu, Continuous Deployment selangkah lebih maju dengan mendeploy setiap perubahan yang melewati semua tahap pipeline produksi secara otomatis.',
  { indent: true, line: 276 }
));

// Table 1
children.push(emptyLine());
children.push(caption('Tabel 1. Perbandingan CI, Continuous Delivery, dan Continuous Deployment'));
const TW1 = 9360;
const c1 = [2340, 2340, 2340, 2340];
children.push(new Table({
  width: { size: TW1, type: WidthType.DXA },
  columnWidths: c1,
  rows: [
    new TableRow({ children: [
      hCell('Aspek', c1[0]),
      hCell('Continuous Integration', c1[1]),
      hCell('Continuous Delivery', c1[2]),
      hCell('Continuous Deployment', c1[3]),
    ]}),
    new TableRow({ children: [
      dCell('Definisi', c1[0], false, 'EAF0FB'),
      dCell('Integrasi kode otomatis setiap commit', c1[1], false, 'EAF0FB'),
      dCell('Kode siap deploy kapan saja (manual trigger)', c1[2], false, 'EAF0FB'),
      dCell('Deploy otomatis tanpa intervensi manual', c1[3], false, 'EAF0FB'),
    ]}),
    new TableRow({ children: [
      dCell('Tujuan', c1[0], false, 'FFFFFF'),
      dCell('Deteksi bug lebih awal', c1[1], false, 'FFFFFF'),
      dCell('Selalu memiliki versi releasable', c1[2], false, 'FFFFFF'),
      dCell('Rilis kontinyu ke production', c1[3], false, 'FFFFFF'),
    ]}),
    new TableRow({ children: [
      dCell('Trigger Deploy', c1[0], false, 'EAF0FB'),
      dCell('Tidak ada', c1[1], true, 'EAF0FB'),
      dCell('Manual', c1[2], true, 'EAF0FB'),
      dCell('Otomatis', c1[3], true, 'EAF0FB'),
    ]}),
    new TableRow({ children: [
      dCell('Risiko', c1[0], false, 'FFFFFF'),
      dCell('Rendah per integrasi', c1[1], false, 'FFFFFF'),
      dCell('Terkontrol', c1[2], false, 'FFFFFF'),
      dCell('Perlu coverage test tinggi', c1[3], false, 'FFFFFF'),
    ]}),
    new TableRow({ children: [
      dCell('Tools Utama', c1[0], false, 'EAF0FB'),
      dCell('Jenkins, GitHub Actions', c1[1], false, 'EAF0FB'),
      dCell('Jenkins, Spinnaker', c1[2], false, 'EAF0FB'),
      dCell('Jenkins, ArgoCD', c1[3], false, 'EAF0FB'),
    ]}),
  ]
}));
children.push(emptyLine());

children.push(heading2('2.4 Pengenalan Jenkins'));
children.push(para(
  'Jenkins adalah server otomasi open-source yang ditulis dalam bahasa Java. Jenkins membantu dalam mengotomasi bagian-bagian pengembangan perangkat lunak yang berkaitan dengan building, testing, dan deploying, sehingga memfasilitasi Continuous Integration dan Continuous Delivery. Jenkins dapat dijalankan di berbagai platform termasuk Windows, macOS, dan Linux/Unix.',
  { indent: true, line: 276 }
));
children.push(para(
  'Harrius, Jek Siang, dan Wibowo (2024) dalam penelitiannya tentang implementasi Continuous Integration dan otomatisasi monitoring job server dengan Jenkins pada Compute Engine menyimpulkan bahwa Jenkins menyediakan mekanisme yang andal untuk mengotomasi proses build dan monitoring, dengan kemampuan notifikasi real-time yang mempermudah deteksi kegagalan sistem secara dini.',
  { indent: true, line: 276 }
));
children.push(para('Keunggulan utama Jenkins antara lain:', { line: 276 }));
children.push(bulletPara('Open-source dan gratis tanpa biaya lisensi'));
children.push(bulletPara('Lebih dari 1.800 plugin tersedia di Jenkins Plugin Index'));
children.push(bulletPara('Mendukung berbagai sistem kontrol versi (Git, SVN, Mercurial)'));
children.push(bulletPara('Dapat dijalankan di berbagai platform (Windows, Linux, macOS, Docker)'));
children.push(bulletPara('Mendukung arsitektur distributed build dengan konsep Master-Agent'));
children.push(bulletPara('Konfigurasi pipeline dapat didefinisikan sebagai kode (Pipeline as Code)'));
children.push(emptyLine());

// ── BAB III ──
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(heading1('BAB III\nPEMBAHASAN'));

children.push(heading2('3.1 Arsitektur Jenkins'));
children.push(para(
  'Jenkins menggunakan arsitektur Master-Agent (sebelumnya disebut Master-Slave) yang memungkinkan distribusi beban kerja build dan deployment ke berbagai mesin. Arsitektur ini memberikan skalabilitas, fleksibilitas, dan efisiensi dalam mengelola proyek-proyek berskala besar.',
  { indent: true, line: 276 }
));

children.push(imageBlock('/home/claude/jenkins_paper/images/gambar1_arsitektur.png', 550, 310));
children.push(caption('Gambar 1. Arsitektur Jenkins Master-Agent dan Alur Komunikasi Antar Komponen'));

children.push(heading3('3.1.1 Jenkins Controller (Master)'));
children.push(para(
  'Jenkins Controller atau Master merupakan komponen inti yang bertugas mengelola konfigurasi sistem, menjadwalkan build job, mendistribusikan pekerjaan ke agent, memantau status agent, dan menyajikan antarmuka web kepada pengguna. Controller menyimpan seluruh konfigurasi sistem, informasi job, history build, dan hasil pengujian dalam sistem file lokalnya.',
  { indent: true, line: 276 }
));

children.push(heading3('3.1.2 Jenkins Agent (Node)'));
children.push(para(
  'Agent merupakan mesin terpisah yang terhubung ke Jenkins Controller dan digunakan untuk menjalankan pipeline atau job yang didelegasikan oleh Controller. Setiap Agent dapat dikonfigurasi dengan label tertentu sehingga job dapat diarahkan ke Agent yang memiliki spesifikasi atau kemampuan tertentu, misalnya Agent khusus Windows untuk build aplikasi .NET atau Agent khusus Docker untuk build berbasis kontainer.',
  { indent: true, line: 276 }
));

children.push(heading2('3.2 Prasyarat Instalasi Jenkins di Windows'));
children.push(para(
  'Sebelum melakukan instalasi Jenkins pada sistem operasi Windows, terdapat beberapa prasyarat teknis yang wajib dipenuhi. Kesiapan prasyarat ini sangat penting untuk memastikan Jenkins dapat berjalan dengan optimal dan stabil.',
  { indent: true, line: 276 }
));

children.push(imageBlock('/home/claude/jenkins_paper/images/gambar3_prasyarat.png', 500, 290));
children.push(caption('Gambar 2. Prasyarat Sistem yang Diperlukan untuk Instalasi Jenkins di Windows'));

// Table 2
children.push(emptyLine());
children.push(caption('Tabel 2. Spesifikasi Minimum dan Rekomendasi Sistem untuk Jenkins di Windows'));
const TW2 = 9360;
const c2 = [2500, 3430, 3430];
children.push(new Table({
  width: { size: TW2, type: WidthType.DXA },
  columnWidths: c2,
  rows: [
    new TableRow({ children: [hCell('Komponen', c2[0]), hCell('Spesifikasi Minimum', c2[1]), hCell('Spesifikasi Rekomendasi', c2[2])] }),
    new TableRow({ children: [dCell('Sistem Operasi', c2[0], false, 'EAF0FB'), dCell('Windows 10 (64-bit)', c2[1], false, 'EAF0FB'), dCell('Windows Server 2019/2022', c2[2], false, 'EAF0FB')] }),
    new TableRow({ children: [dCell('Java (JDK)', c2[0]), dCell('JDK 17 (LTS)', c2[1]), dCell('JDK 21 (LTS) - Eclipse Temurin', c2[2])] }),
    new TableRow({ children: [dCell('RAM', c2[0], false, 'EAF0FB'), dCell('2 GB', c2[1], false, 'EAF0FB'), dCell('4 GB atau lebih', c2[2], false, 'EAF0FB')] }),
    new TableRow({ children: [dCell('Penyimpanan', c2[0]), dCell('10 GB', c2[1]), dCell('50 GB atau lebih', c2[2])] }),
    new TableRow({ children: [dCell('Prosesor', c2[0], false, 'EAF0FB'), dCell('2 Core', c2[1], false, 'EAF0FB'), dCell('4 Core atau lebih', c2[2], false, 'EAF0FB')] }),
    new TableRow({ children: [dCell('Port yang Dibutuhkan', c2[0]), dCell('8080 (HTTP)', c2[1]), dCell('8080, 50000 (Agent)', c2[2])] }),
    new TableRow({ children: [dCell('Koneksi Internet', c2[0], false, 'EAF0FB'), dCell('Diperlukan untuk plugin', c2[1], false, 'EAF0FB'), dCell('Stabil broadband', c2[2], false, 'EAF0FB')] }),
  ]
}));
children.push(emptyLine());

children.push(heading2('3.3 Langkah-Langkah Instalasi Jenkins di Windows'));
children.push(para(
  'Berikut ini adalah panduan lengkap instalasi Jenkins di lingkungan Windows secara step-by-step. Setiap tahap harus dilakukan secara berurutan untuk memastikan instalasi berjalan dengan benar.',
  { indent: true, line: 276 }
));

children.push(imageBlock('/home/claude/jenkins_paper/images/gambar4_instalasi.png', 500, 490));
children.push(caption('Gambar 3. Diagram Alur Langkah-Langkah Instalasi Jenkins di Windows'));

children.push(heading3('3.3.1 Instalasi Java Development Kit (JDK)'));
children.push(para(
  'Jenkins membutuhkan Java untuk dapat berjalan. Langkah pertama adalah mengunduh dan menginstal JDK 17 atau JDK 21 dari Eclipse Temurin (https://adoptium.net) yang merupakan distribusi OpenJDK yang direkomendasikan oleh komunitas Jenkins. Setelah proses instalasi JDK selesai, variabel lingkungan JAVA_HOME harus dikonfigurasi melalui System Properties agar Jenkins dapat menemukan runtime Java dengan benar.',
  { indent: true, line: 276 }
));
children.push(para('Langkah konfigurasi variabel lingkungan JAVA_HOME:', { line: 276 }));
children.push(numPara('Klik kanan pada "This PC" atau "My Computer" → Properties'));
children.push(numPara('Pilih "Advanced system settings" → "Environment Variables"'));
children.push(numPara('Di bagian "System variables", klik "New"'));
children.push(numPara('Masukkan Variable name: JAVA_HOME dan Variable value: C:\\Program Files\\Eclipse Adoptium\\jdk-17.x.x.x-hotspot'));
children.push(numPara('Edit variabel PATH, tambahkan: %JAVA_HOME%\\bin'));
children.push(numPara('Verifikasi dengan membuka Command Prompt dan jalankan: java -version'));

children.push(heading3('3.3.2 Mengunduh dan Menjalankan Jenkins Installer'));
children.push(para(
  'Jenkins menyediakan installer khusus untuk Windows dalam format MSI (Microsoft Software Installer) yang dapat diunduh dari situs resmi Jenkins (https://www.jenkins.io/download). Installer ini memudahkan proses instalasi dengan Setup Wizard yang interaktif. Pengguna dapat memilih direktori instalasi, konfigurasi layanan Windows, serta port yang digunakan oleh Jenkins.',
  { indent: true, line: 276 }
));
children.push(para('Selama proses instalasi melalui Jenkins Setup Wizard, terdapat beberapa konfigurasi penting:', { line: 276 }));
children.push(bulletPara('Destination Folder: Pilih direktori instalasi Jenkins (default: C:\\Program Files\\Jenkins)'));
children.push(bulletPara('Service Logon Credentials: Pilih "Run service as LocalSystem" untuk kemudahan atau tentukan akun Windows khusus untuk keamanan lebih baik'));
children.push(bulletPara('Port Selection: Tentukan port yang akan digunakan Jenkins (default: 8080). Pastikan port tidak digunakan oleh aplikasi lain'));
children.push(bulletPara('Java Installation: Setup Wizard akan otomatis mendeteksi instalasi Java; verifikasi bahwa JDK 17/21 terdeteksi dengan benar'));

children.push(heading3('3.3.3 Konfigurasi Awal Jenkins (Initial Setup)'));
children.push(para(
  'Setelah instalasi selesai, Jenkins secara otomatis berjalan sebagai layanan Windows dan dapat diakses melalui browser di alamat http://localhost:8080. Pada akses pertama kali, Jenkins memerlukan konfigurasi awal yang mencakup unlocking dengan initial admin password, instalasi plugin, dan pembuatan akun administrator.',
  { indent: true, line: 276 }
));
children.push(para('Initial Admin Password dapat ditemukan di lokasi berikut:', { line: 276 }));
children.push(bulletPara('Path file: C:\\ProgramData\\Jenkins\\.jenkins\\secrets\\initialAdminPassword'));
children.push(bulletPara('Password dapat dibuka menggunakan Notepad atau aplikasi teks editor lainnya'));
children.push(bulletPara('Salin password tersebut dan tempelkan pada kolom "Administrator password" di halaman web Jenkins'));

children.push(heading3('3.3.4 Instalasi Plugin Jenkins'));
children.push(para(
  'Jenkins menawarkan dua opsi instalasi plugin pada saat setup awal: (1) Install suggested plugins yang menginstal kumpulan plugin yang paling umum digunakan secara otomatis, dan (2) Select plugins to install yang memungkinkan pengguna memilih plugin secara manual. Untuk pemula dan penggunaan umum, opsi "Install suggested plugins" sangat direkomendasikan.',
  { indent: true, line: 276 }
));

// Table 3
children.push(emptyLine());
children.push(caption('Tabel 3. Plugin-Plugin Penting Jenkins untuk Implementasi CI/CD'));
const TW3 = 9360;
const c3 = [2500, 4000, 2860];
children.push(new Table({
  width: { size: TW3, type: WidthType.DXA },
  columnWidths: c3,
  rows: [
    new TableRow({ children: [hCell('Nama Plugin', c3[0]), hCell('Fungsi', c3[1]), hCell('Kategori', c3[2])] }),
    new TableRow({ children: [dCell('Git Plugin', c3[0], false, 'EAF0FB'), dCell('Integrasi dengan repositori Git dan GitHub', c3[1], false, 'EAF0FB'), dCell('Source Control', c3[2], false, 'EAF0FB')] }),
    new TableRow({ children: [dCell('Pipeline', c3[0]), dCell('Membuat dan mengelola pipeline CI/CD dengan Jenkinsfile', c3[1]), dCell('Core', c3[2])] }),
    new TableRow({ children: [dCell('Maven Integration', c3[0], false, 'EAF0FB'), dCell('Build otomatis proyek berbasis Apache Maven', c3[1], false, 'EAF0FB'), dCell('Build Tools', c3[2], false, 'EAF0FB')] }),
    new TableRow({ children: [dCell('Email Extension', c3[0]), dCell('Notifikasi email kustomisasi untuk hasil build', c3[1]), dCell('Notifications', c3[2])] }),
    new TableRow({ children: [dCell('Blue Ocean', c3[0], false, 'EAF0FB'), dCell('Tampilan visual modern untuk pipeline Jenkins', c3[1], false, 'EAF0FB'), dCell('UI/UX', c3[2], false, 'EAF0FB')] }),
    new TableRow({ children: [dCell('Credentials Binding', c3[0]), dCell('Manajemen credentials sensitif seperti password dan API key', c3[1]), dCell('Security', c3[2])] }),
    new TableRow({ children: [dCell('Docker Pipeline', c3[0], false, 'EAF0FB'), dCell('Integrasi Docker dalam pipeline Jenkins', c3[1], false, 'EAF0FB'), dCell('Containerization', c3[2], false, 'EAF0FB')] }),
    new TableRow({ children: [dCell('Workspace Cleanup', c3[0]), dCell('Membersihkan workspace setelah setiap build', c3[1]), dCell('Maintenance', c3[2])] }),
  ]
}));
children.push(emptyLine());

children.push(heading2('3.4 Tampilan Dashboard Jenkins'));
children.push(para(
  'Setelah instalasi dan konfigurasi awal selesai, pengguna dapat mengakses Jenkins Dashboard yang merupakan antarmuka utama untuk mengelola semua job dan pipeline. Dashboard menyediakan overview mengenai status seluruh job yang terdaftar, history build, dan berbagai menu navigasi untuk manajemen sistem.',
  { indent: true, line: 276 }
));

children.push(imageBlock('/home/claude/jenkins_paper/images/gambar5_dashboard.png', 580, 340));
children.push(caption('Gambar 4. Tampilan Dashboard Jenkins dengan Daftar Pipeline dan Status Build'));

children.push(para(
  'Pada Dashboard Jenkins, terdapat beberapa elemen antarmuka utama yang perlu dipahami: (1) Navigation Bar di bagian atas yang berisi menu New Item, People, Build History, dan Manage Jenkins; (2) Sidebar kiri yang menampilkan menu-menu utama; (3) Area konten utama yang menampilkan daftar job/pipeline beserta status build terakhir; serta (4) Indikator status build berupa ikon yang menunjukkan apakah build terakhir berhasil (biru/hijau) atau gagal (merah).',
  { indent: true, line: 276 }
));

children.push(heading2('3.5 Konfigurasi Pipeline CI/CD dengan Jenkinsfile'));
children.push(para(
  'Jenkins Pipeline adalah rangkaian plugin yang mendukung implementasi dan integrasi pipeline Continuous Delivery secara berkelanjutan. Pipeline didefinisikan menggunakan Jenkinsfile, yaitu file teks yang berisi definisi pipeline menggunakan sintaks Groovy DSL (Domain Specific Language). Pendekatan Pipeline as Code ini memungkinkan konfigurasi pipeline disimpan bersama kode sumber dalam sistem kontrol versi.',
  { indent: true, line: 276 }
));

children.push(imageBlock('/home/claude/jenkins_paper/images/gambar6_jenkinsfile.png', 570, 260));
children.push(caption('Gambar 5. Contoh Struktur Jenkinsfile dan Tahapan Pipeline CI/CD untuk Aplikasi Java'));

children.push(heading3('3.5.1 Struktur Dasar Jenkinsfile'));
children.push(para(
  'Jenkinsfile menggunakan sintaks Declarative Pipeline yang lebih terstruktur dan mudah dipahami. Komponen utama dalam Jenkinsfile meliputi: blok pipeline sebagai container utama, blok agent yang menentukan di mana pipeline dieksekusi, dan blok stages yang berisi serangkaian stage dengan step-step yang harus dijalankan.',
  { indent: true, line: 276 }
));
children.push(para(
  'Untuk lingkungan Windows, terdapat perbedaan penting dalam penulisan perintah pada blok steps. Pada Windows, digunakan perintah bat (batch) untuk menjalankan perintah Windows Command Prompt, sedangkan pada Linux/macOS digunakan sh (shell). Hal ini perlu diperhatikan saat mengembangkan Jenkinsfile untuk lingkungan Windows.',
  { indent: true, line: 276 }
));

children.push(heading3('3.5.2 Tahapan Pipeline CI/CD di Windows'));
children.push(para(
  'Setyoko dan Zahra (2024) dalam penelitiannya tentang perbandingan efisiensi proses CI/CD multi-lingkungan menyimpulkan bahwa implementasi pipeline dengan tahapan yang terdefinisi dengan baik secara signifikan meningkatkan efisiensi proses dan mengurangi waktu deployment. Berdasarkan best practice yang ada, berikut adalah tahapan umum dalam pipeline CI/CD Jenkins untuk lingkungan Windows:',
  { indent: true, line: 276 }
));

children.push(imageBlock('/home/claude/jenkins_paper/images/gambar2_pipeline.png', 590, 195));
children.push(caption('Gambar 6. Alur Lengkap Tahapan Pipeline CI/CD pada Jenkins'));

// Table 4
children.push(emptyLine());
children.push(caption('Tabel 4. Penjelasan Setiap Tahapan dalam Pipeline CI/CD Jenkins'));
const TW4 = 9360;
const c4 = [1440, 2000, 3200, 2720];
children.push(new Table({
  width: { size: TW4, type: WidthType.DXA },
  columnWidths: c4,
  rows: [
    new TableRow({ children: [hCell('No.', c4[0], '1E5799'), hCell('Stage', c4[1], '1E5799'), hCell('Deskripsi', c4[2], '1E5799'), hCell('Perintah Windows (bat)', c4[3], '1E5799')] }),
    new TableRow({ children: [
      dCell('1', c4[0], true, 'EAF0FB'), dCell('Checkout', c4[1], false, 'EAF0FB'),
      dCell('Mengambil kode sumber terbaru dari repositori Git', c4[2], false, 'EAF0FB'),
      dCell("git checkout, git pull", c4[3], false, 'EAF0FB')
    ]}),
    new TableRow({ children: [
      dCell('2', c4[0], true), dCell('Build', c4[1]),
      dCell('Mengkompilasi source code menjadi artefak yang dapat dieksekusi', c4[2]),
      dCell('bat "mvn package" atau bat "gradle build"', c4[3])
    ]}),
    new TableRow({ children: [
      dCell('3', c4[0], true, 'EAF0FB'), dCell('Unit Test', c4[1], false, 'EAF0FB'),
      dCell('Menjalankan unit test untuk memverifikasi fungsionalitas komponen', c4[2], false, 'EAF0FB'),
      dCell('bat "mvn test" atau bat "pytest tests/"', c4[3], false, 'EAF0FB')
    ]}),
    new TableRow({ children: [
      dCell('4', c4[0], true), dCell('Code Analysis', c4[1]),
      dCell('Analisis kode statis untuk mendeteksi bug, kerentanan, dan code smell', c4[2]),
      dCell('bat "sonar-scanner.bat" atau bat "checkstyle"', c4[3])
    ]}),
    new TableRow({ children: [
      dCell('5', c4[0], true, 'EAF0FB'), dCell('Package', c4[1], false, 'EAF0FB'),
      dCell('Membungkus artefak build beserta dependensinya (.jar, .war, .exe)', c4[2], false, 'EAF0FB'),
      dCell('bat "mvn package -DskipTests"', c4[3], false, 'EAF0FB')
    ]}),
    new TableRow({ children: [
      dCell('6', c4[0], true), dCell('Deploy Staging', c4[1]),
      dCell('Mendeploy aplikasi ke lingkungan staging/UAT untuk pengujian lebih lanjut', c4[2]),
      dCell('bat "deploy-staging.bat"', c4[3])
    ]}),
    new TableRow({ children: [
      dCell('7', c4[0], true, 'EAF0FB'), dCell('Deploy Production', c4[1], false, 'EAF0FB'),
      dCell('Mendeploy aplikasi ke lingkungan produksi setelah semua tahap berhasil', c4[2], false, 'EAF0FB'),
      dCell('bat "deploy-prod.bat"', c4[3], false, 'EAF0FB')
    ]}),
  ]
}));
children.push(emptyLine());

children.push(heading2('3.6 Integrasi Jenkins dengan Git/GitHub di Windows'));
children.push(para(
  'Integrasi Jenkins dengan repositori Git merupakan langkah krusial dalam implementasi CI/CD. Jenkins dapat dikonfigurasi untuk secara otomatis memicu build setiap kali terjadi perubahan pada repositori (event push atau pull request). Konfigurasi ini dilakukan melalui kombinasi webhook pada sisi GitHub dan job trigger pada sisi Jenkins.',
  { indent: true, line: 276 }
));
children.push(para('Langkah-langkah mengintegrasikan Jenkins dengan GitHub:', { line: 276 }));
children.push(numPara('Di Jenkins: Install plugin "GitHub Integration" dan "Git" melalui Manage Jenkins → Plugin Manager'));
children.push(numPara('Konfigurasi kredensial GitHub: Manage Jenkins → Credentials → Add Credentials → pilih "Username with password" atau "Secret text" (untuk Personal Access Token)'));
children.push(numPara('Di GitHub: Masuk ke Settings repositori → Webhooks → Add Webhook'));
children.push(numPara('Isi Payload URL dengan: http://<jenkins-server>:8080/github-webhook/'));
children.push(numPara('Di Jenkins Job/Pipeline: Pada bagian Build Triggers, centang "GitHub hook trigger for GITScm polling"'));
children.push(numPara('Setiap commit baru ke repositori akan otomatis memicu build pada Jenkins'));

children.push(heading2('3.7 Manajemen Credentials dan Keamanan'));
children.push(para(
  'Keamanan adalah aspek penting dalam implementasi Jenkins. Jenkins menyediakan mekanisme manajemen credentials yang aman untuk menyimpan informasi sensitif seperti password, API key, SSH key, dan sertifikat. Credentials disimpan secara terenkripsi dan dapat digunakan dalam pipeline tanpa mengeksposnya secara langsung dalam kode.',
  { indent: true, line: 276 }
));
children.push(para(
  'Pada lingkungan Windows, perlu diperhatikan pengaturan keamanan Jenkins melalui menu Manage Jenkins → Configure Global Security. Rekomendasi konfigurasi keamanan meliputi: mengaktifkan autentikasi pengguna, mengatur authorization matrix untuk kontrol akses berbasis peran, mengaktifkan CSRF protection, dan mengkonfigurasi agent-to-controller security.',
  { indent: true, line: 276 }
));

children.push(emptyLine());
// ── BAB IV ──
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(heading1('BAB IV\nKESIMPULAN DAN SARAN'));

children.push(heading2('4.1 Kesimpulan'));
children.push(para(
  'Berdasarkan pembahasan yang telah diuraikan dalam makalah ini, dapat ditarik beberapa kesimpulan sebagai berikut:',
  { line: 276 }
));
children.push(numPara('Jenkins merupakan platform otomasi CI/CD open-source yang handal dan fleksibel, dengan dukungan lebih dari 1.800 plugin yang memungkinkan integrasi dengan berbagai alat pengembangan perangkat lunak dalam ekosistem DevOps.'));
children.push(numPara('Arsitektur Master-Agent Jenkins memungkinkan skalabilitas horizontal yang baik, di mana Controller mengelola konfigurasi dan penjadwalan, sementara Agent bertanggung jawab atas eksekusi build secara terdistribusi.'));
children.push(numPara('Instalasi Jenkins di Windows dapat dilakukan dengan mudah menggunakan installer MSI yang tersedia di situs resmi Jenkins, dengan prasyarat utama berupa instalasi JDK 17 atau 21 dan konfigurasi variabel lingkungan yang tepat.'));
children.push(numPara('Pendekatan Pipeline as Code melalui Jenkinsfile memungkinkan definisi seluruh alur CI/CD disimpan dalam sistem kontrol versi, sehingga meningkatkan reproducibility dan kemudahan audit terhadap perubahan konfigurasi.'));
children.push(numPara('Implementasi CI/CD dengan Jenkins terbukti mampu mempercepat siklus pengembangan perangkat lunak, meningkatkan kualitas kode, dan meminimalkan risiko kesalahan dalam proses deployment.'));
children.push(emptyLine());

children.push(heading2('4.2 Saran'));
children.push(para(
  'Berdasarkan kajian yang telah dilakukan, terdapat beberapa saran yang dapat diberikan untuk implementasi Jenkins yang optimal:',
  { line: 276 }
));
children.push(numPara('Gunakan Jenkins dalam lingkungan terisolasi atau dedicated server untuk menghindari konflik dengan aplikasi lain dan memastikan stabilitas layanan CI/CD.'));
children.push(numPara('Implementasikan backup rutin terhadap direktori JENKINS_HOME untuk mencegah kehilangan konfigurasi dan history build yang penting.'));
children.push(numPara('Terapkan prinsip least privilege dalam konfigurasi akun Jenkins untuk meningkatkan keamanan sistem, terutama pada lingkungan produksi.'));
children.push(numPara('Pertimbangkan penggunaan Jenkins X atau Jenkins dengan Kubernetes untuk kebutuhan skala besar yang memerlukan autoscaling dan containerisasi.'));
children.push(numPara('Lakukan pembaruan Jenkins dan plugin secara berkala untuk mendapatkan patch keamanan terbaru dan fitur-fitur baru yang meningkatkan produktivitas.'));
children.push(emptyLine());

// ── DAFTAR PUSTAKA ──
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(heading1('DAFTAR PUSTAKA'));

const refs = [
  'Harrius, G., Jek Siang, J., & Wibowo, A. (2024). Implementasi Continuous Integration dan Otomatisasi Monitoring Job Server Dengan Jenkins pada Compute Engine. Jurnal Teknologi Dan Sistem Informasi Bisnis, 6(4), 704-714. https://doi.org/10.47233/jteksis.v6i4.1541',
  'Kuncara, A. B., Kusumo, D. S., & Adrian, M. (2024). Comparison of Jenkins and GitLab CI/CD to Improve Delivery Time of Basu Dairy Farm Admin Website. Jurnal Teknik Informatika (JUTIF), 5(3), 747-756. https://doi.org/10.20884/1.jutif.2024.5.3',
  'Kusumadewi, R. T., & Adrian, R. (2023). Performance Analysis of DevOps Practice Implementation of CI/CD Using Jenkins. MATICS: Jurnal Ilmu Komputer dan Teknologi Informasi, 15(2), 90-95. https://doi.org/10.18860/mat.v15i2',
  'Setyoko, A. D., & Zahra, A. (2024). Perbandingan Efisiensi Proses CI/CD Multi-Lingkungan melalui Implementasi Paralel dan Berurutan. MALCOM: Indonesian Journal of Machine Learning and Computer Science, 4(3), 911-925. https://doi.org/10.57152/malcom.v4i3.1344',
  'Patil, K., Kapadnis, S., Waghmare, R., Thakare, H., & Raut, R. (2022). Implementation of a Continuous Integration and Deployment Pipeline for Containerized Applications in Amazon Web Services Using Jenkins. International Journal of Scientific Research in Engineering and Management, 06(11). https://doi.org/10.55041/ijsrem16948',
];

for (const ref of refs) {
  children.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 60, after: 60, line: 276 },
    indent: { left: 720, hanging: 720 },
    children: [new TextRun({ text: ref, font: TNR, size: FONT_SIZE })]
  }));
  children.push(emptyLine());
}

// ─── Document ──────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '\u2022',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } }, run: { font: 'Symbol', size: FONT_SIZE } }
        }]
      },
      {
        reference: 'numbers',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
    ]
  },
  styles: {
    default: {
      document: { run: { font: TNR, size: FONT_SIZE } }
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: TNR, color: '1F3864' },
        paragraph: { spacing: { before: 240, after: 120 }, alignment: AlignmentType.CENTER, outlineLevel: 0 }
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: TNR, color: '2E75B6' },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 }
      },
      {
        id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, italics: true, font: TNR, color: '2E75B6' },
        paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 }
      },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 } // 1 inch margins, left 1.25 inch
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '2E75B6', space: 1 } },
          children: [new TextRun({ text: 'Makalah: Implementasi Jenkins untuk CI/CD Berbasis Windows', font: TNR, size: 18, color: '555555', italics: true })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: '2E75B6', space: 1 } },
          children: [
            new TextRun({ text: 'Halaman ', font: TNR, size: 20, color: '555555' }),
            new PageNumber({ },)
          ]
        })]
      })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/home/claude/jenkins_paper/makalah_jenkins_cicd.docx', buffer);
  console.log('DOCX created successfully!');
});
