/**
 * LAPORAN BULANAN via email: berwarna, dengan diagram + perbandingan 12 bulan.
 * Simpan sebagai file "Laporan" di proyek Apps Script yang sama dengan Code.gs.
 */
const EMAIL_TUJUAN = ''; // kosongkan = kirim ke email akun Google kamu
const BLN_ID = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const BLN_PDK = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];

// Jalankan SEKALI: memasang pengecekan harian jam 20.00-21.00
function pasangJadwal() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'cekAkhirBulan').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('cekAkhirBulan').timeBased().everyDays(1).atHour(20).create();
}

// Setiap hari dicek: kalau besok tanggal 1, hari ini akhir bulan -> kirim laporan
function cekAkhirBulan() {
  const tz = SpreadsheetApp.openById(SHEET_ID).getSpreadsheetTimeZone();
  if (Utilities.formatDate(new Date(Date.now() + 864e5), tz, 'd') === '1') kirimLaporanBulanan();
}

// Bisa dijalankan manual untuk uji coba. Contoh bulan tertentu: kirimLaporanBulanan('2026-10')
function kirimLaporanBulanan(ym) {
  const ss = SpreadsheetApp.openById(SHEET_ID), tz = ss.getSpreadsheetTimeZone(), now = new Date();
  const cur = Utilities.formatDate(now, tz, 'yyyy-MM');
  ym = ym || cur;
  const [y, m] = ym.split('-').map(Number);
  const ms = [];
  for (let k = 11; k >= 0; k--) ms.push(Utilities.formatDate(new Date(y, m - 1 - k, 15), tz, 'yyyy-MM'));
  const lbl = k => BLN_PDK[+k.slice(5) - 1] + ' ' + k.slice(2, 4);
  const rp = n => (n < 0 ? '-' : '') + 'Rp ' + Math.round(Math.abs(n)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');

  // ---- data 12 bulan ----
  const all = readTx(ss, 'in').concat(readTx(ss, 'out')), A = {};
  ms.forEach(k => A[k] = { i: 0, o: 0 });
  all.forEach(x => { const a = A[x.d.slice(0, 7)]; if (a) a[x.t === 'in' ? 'i' : 'o'] += x.n; });
  const c = A[ym], prev = A[ms[10]];
  const pembagi = ym === cur ? Number(Utilities.formatDate(now, tz, 'd')) : new Date(y, m, 0).getDate();
  const lalu = ms.filter(k => k !== ym && A[k].o > 0);
  const avgO = lalu.length ? lalu.reduce((s, k) => s + A[k].o, 0) / lalu.length : 0;
  const ada = ms.filter(k => A[k].o > 0);
  const boros = ada.reduce((a, k) => !a || A[k].o > A[a].o ? k : a, null);
  const hemat = ada.reduce((a, k) => !a || A[k].o < A[a].o ? k : a, null);
  const pct = (a, b) => b ? Math.round((a - b) / b * 100) : null;
  const pill = d => d === null ? '<span style="color:#64748b">belum ada pembanding</span>' :
    d > 0 ? `<b style="color:#e11d48">▲ ${d}% lebih boros</b>` : d < 0 ? `<b style="color:#0d9488">▼ ${-d}% lebih hemat</b>` : '<b>sama</b>';

  // ---- pengeluaran bulan ini per kategori & grup ----
  const setup = readSetup(ss), grp = k => k.replace(/\s*week\s*\d+$/i, '');
  const gOf = k => { const g = setup.out.find(g => g[1].includes(k)); return g ? g[0] : 'Lainnya'; };
  const byCat = {}, byGrp = {};
  all.filter(x => x.d.slice(0, 7) === ym && x.t === 'out').forEach(x => {
    const k = grp(x.k); byCat[k] = (byCat[k] || 0) + x.n;
    const g = gOf(x.k); byGrp[g] = (byGrp[g] || 0) + x.n;
  });
  const top = Object.entries(byCat).sort((a, b) => b[1] - a[1]).slice(0, 5);
  const grupArr = Object.entries(byGrp).filter(e => e[1] > 0);

  // ---- diagram (gambar PNG, dibuat oleh Apps Script) ----
  const juta = n => Math.round(n / 1e4) / 100;
  const tabel = (kolom, baris) => {
    const t = Charts.newDataTable();
    kolom.forEach((n, j) => t.addColumn(j ? Charts.ColumnType.NUMBER : Charts.ColumnType.STRING, n));
    baris.forEach(r => t.addRow(r));
    return t.build();
  };
  const gambar = (nama, f) => { try { return f().build().getBlob().setName(nama + '.png'); } catch (e) { return null; } };
  const WARNA = { Needs: '#2563eb', Wants: '#f59e0b', Saving: '#0d9488', Lainnya: '#94a3b8' };
  const blobs = {
    tren: gambar('tren', () => Charts.newColumnChart()
      .setDataTable(tabel(['Bulan', 'Pemasukan', 'Pengeluaran'], ms.map(k => [lbl(k), juta(A[k].i), juta(A[k].o)])))
      .setColors(['#0d9488', '#7c3aed']).setDimensions(600, 300).setLegendPosition(Charts.Position.BOTTOM)
      .setTitle('Pemasukan vs pengeluaran, 12 bulan (juta Rp)')),
    saldo: gambar('saldo', () => Charts.newLineChart()
      .setDataTable(tabel(['Bulan', 'Saldo bersih'], ms.map(k => [lbl(k), juta(A[k].i - A[k].o)])))
      .setColors(['#2563eb']).setDimensions(600, 260).setLegendPosition(Charts.Position.NONE)
      .setTitle('Saldo bersih per bulan (juta Rp)')),
    donat: gambar('donat', () => Charts.newPieChart()
      .setDataTable(tabel(['Grup', 'Rp'], grupArr.map(e => [e[0], e[1]])))
      .setColors(grupArr.map(e => WARNA[e[0]])).setOption('pieHole', 0.5).setDimensions(600, 260)
      .setLegendPosition(Charts.Position.RIGHT).setTitle('Komposisi pengeluaran: Needs / Wants / Saving'))
  };
  const inline = {};
  Object.keys(blobs).forEach(k => { if (blobs[k]) inline[k] = blobs[k]; });
  const im = k => blobs[k] ? `<img src="cid:${k}" width="100%" style="display:block;border-radius:12px;margin:8px 0">` : '';

  // ---- tampilan email ----
  const card = (bg, judul, nilai) => `<td width="33%" style="padding:4px"><div style="background:${bg};border-radius:14px;padding:12px;color:#fff"><div style="font-size:11px;opacity:.9">${judul}</div><div style="font-size:14px;font-weight:700;margin-top:4px">${nilai}</div></div></td>`;
  const box = (judul, isi) => `<div style="background:#fff;border:1px solid #e8ecf4;border-radius:16px;padding:14px;margin:12px 0"><div style="font-weight:700;font-size:15px;margin-bottom:6px">${judul}</div>${isi}</div>`;
  const info = (ikon, teks) => `<div style="padding:7px 0;font-size:14px">${ikon} ${teks}</div>`;
  const PAL = ['#7c3aed', '#2563eb', '#0d9488', '#f59e0b', '#e11d48'], max = top.length ? top[0][1] : 1;
  const bars = top.length ? top.map(([k, v], j) => `<div style="margin:10px 0 3px;font-size:13px"><b>${k}</b><span style="float:right;color:#64748b">${rp(v)}</span></div><div style="background:#eef2ff;border-radius:6px"><div style="height:10px;width:${Math.max(4, Math.round(v / max * 100))}%;background:${PAL[j]};border-radius:6px"></div></div>`).join('') : '<div style="color:#64748b">Belum ada pengeluaran bulan ini.</div>';
  const baris = ms.filter(k => A[k].i || A[k].o || k === ym).map(k => {
    const s = A[k].i - A[k].o;
    return `<tr style="background:${k === ym ? '#ede9fe' : '#fff'}"><td style="padding:7px 6px"><b>${lbl(k)}</b></td><td align="right" style="color:#0d9488">${rp(A[k].i)}</td><td align="right" style="color:#7c3aed">${rp(A[k].o)}</td><td align="right" style="padding-right:6px;font-weight:700;color:${s < 0 ? '#e11d48' : '#2563eb'}">${rp(s)}</td></tr>`;
  }).join('');
  const insight =
    info('📊', `Pengeluaran vs bulan lalu (${rp(prev.o)}): ${pill(pct(c.o, prev.o))}`) +
    info('📅', `Pengeluaran vs rata-rata bulan sebelumnya (${rp(avgO)}): ${pill(pct(c.o, avgO))}`) +
    info('💸', `Rata-rata pengeluaran harian: <b>${rp(c.o / pembagi)}</b>`) +
    (boros ? info('🔥', `Bulan paling boros: <b>${lbl(boros)}</b> (${rp(A[boros].o)})`) : '') +
    (hemat ? info('🌱', `Bulan paling hemat: <b>${lbl(hemat)}</b> (${rp(A[hemat].o)})`) : '');
  const saldo = c.i - c.o, judul = `${BLN_ID[m - 1]} ${y}`;

  const html = `<div style="background:#f5f7ff;padding:16px 8px;font-family:Arial,Helvetica,sans-serif;color:#0f172a"><div style="max-width:600px;margin:auto">
  <div style="background:#4f46e5;background-image:linear-gradient(135deg,#2563eb,#7c3aed);border-radius:20px;padding:22px;color:#fff">
    <div style="font-size:13px;opacity:.85">Laporan keuangan</div><div style="font-size:26px;font-weight:800;margin:2px 0 6px">${judul}</div>
    <div style="font-size:13px;opacity:.9">Saldo bersih <b style="font-size:20px">${rp(saldo)}</b></div></div>
  <table width="100%" cellspacing="0" style="margin-top:8px"><tr>${card('#0d9488', 'Pemasukan', rp(c.i))}${card('#7c3aed', 'Pengeluaran', rp(c.o))}${card(saldo < 0 ? '#e11d48' : '#2563eb', 'Saldo', rp(saldo))}</tr></table>
  ${box('Insight bulan ini', insight)}
  ${box('Tren 12 bulan', im('tren') + im('saldo'))}
  ${box('Pengeluaran bulan ini', im('donat') + '<div style="font-size:12px;color:#64748b;margin-top:8px">5 kategori terbesar</div>' + bars)}
  ${box('Perbandingan 12 bulan', `<table width="100%" cellspacing="0" style="font-size:12px;border-collapse:collapse"><tr style="color:#64748b"><td style="padding:4px 6px">Bulan</td><td align="right">Masuk</td><td align="right">Keluar</td><td align="right" style="padding-right:6px">Saldo</td></tr>${baris}</table>`)}
  <div style="text-align:center;font-size:11px;color:#94a3b8;margin-top:12px">Dibuat otomatis dari spreadsheet Keuanganku</div></div></div>`;

  MailApp.sendEmail({
    to: EMAIL_TUJUAN || Session.getEffectiveUser().getEmail(),
    subject: 'Laporan Keuangan ' + judul,
    body: 'Laporan keuangan ' + judul + '. Buka email ini di aplikasi email yang mendukung tampilan HTML.',
    htmlBody: html,
    inlineImages: inline
  });
}
