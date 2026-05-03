// Mengambil elemen dari DOM
const myForm = document.getElementById('myForm');
const myInput = document.getElementById('myInput');
const lokasiInput = document.getElementById('lokasiInput');
const resultDiv = document.getElementById('resultDiv');
const averageResult = document.getElementById('averageResult');
const warningNote = document.getElementById('warningNote');

// Array untuk menyimpan semua data
let dataArray = [];

// Counter untuk tracking cell mana yang akan diisi
let cellCounter = 1;

// Menambah event listener saat form di-submit
myForm.addEventListener('submit', function(e) {
    // 1. Mencegah halaman refresh otomatis
    e.preventDefault();

    // 2. Ambil nilai dari input
    const value = myInput.value;

    // 3. Masukkan data ke cell dan array jika belum penuh
    if (cellCounter <= 10) {
        const cellId = 'cell' + cellCounter;
        document.getElementById(cellId).innerText = value;
        dataArray.push(parseFloat(value));
        cellCounter++;
        
        // 4. Jika semua 10 cell sudah terisi, hitung rata-rata
        if (cellCounter === 11) {
            hitungRataRata();
        }
    }

    console.log("Data diterima:", value);

    // 5. Kosongkan kembali kolom input
    myInput.value = "";

    // 6. Kembalikan fokus kursor ke kotak input agar bisa langsung mengetik lagi
    myInput.focus();
});

// Event listener untuk edit cell
const editableCells = document.querySelectorAll('.editable-cell');
editableCells.forEach((cell, index) => {
    cell.addEventListener('click', function() {
        if (this.innerText === '-') {
            return; // Jangan edit jika masih kosong
        }
        
        const currentValue = this.innerText;
        const input = document.createElement('input');
        input.type = 'text';
        input.value = currentValue;
        input.className = 'edit-input';
        
        this.innerText = '';
        this.appendChild(input);
        input.focus();
        input.select();
        
        // Simpan perubahan saat enter atau blur
        const saveEdit = () => {
            const newValue = input.value;
            if (newValue.trim() !== '') {
                this.innerText = newValue;
                const cellNumber = parseInt(this.id.replace('cell', ''));
                dataArray[cellNumber - 1] = parseFloat(newValue);
                
                // Jika semua 10 cell sudah terisi, hitung ulang rata-rata
                if (dataArray.length === 10 && dataArray.every(val => !isNaN(val))) {
                    hitungRataRata();
                }
            } else {
                this.innerText = currentValue;
            }
        };
        
        input.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                saveEdit();
            }
        });
        
        input.addEventListener('blur', saveEdit);
    });
});

// Fungsi untuk menghitung rata-rata
function hitungRataRata() {
    const validData = dataArray.filter(val => !isNaN(val));
    if (validData.length === 10) {
        const total = validData.reduce((sum, num) => sum + num, 0);
        const rataRata = total / 10;
        
        averageResult.innerText = rataRata.toFixed(2) + ' dB';
        
        // Tambahkan warna merah dan catatan jika melebihi 85 dB
        if (rataRata > 85) {
            averageResult.classList.add('high-value');
            warningNote.style.display = 'block';
        } else {
            averageResult.classList.remove('high-value');
            warningNote.style.display = 'none';
        }
        
        resultDiv.style.display = 'block';
    }
}

// Fungsi export ke Excel
function exportToExcel() {
    // Cek apakah lokasi sudah diisi
    if (lokasiInput.value.trim() === '') {
        alert('Mohon isi lokasi terlebih dahulu!');
        lokasiInput.focus();
        return;
    }
    
    // Ambil data dari tabel
    const currentData = [];
    for (let i = 1; i <= 10; i++) {
        currentData.push(document.getElementById('cell' + i).innerText);
    }
    
    // Ambil rata-rata
    const rataRata = averageResult.innerText;
    const lokasi = lokasiInput.value;
    
    // Ambil history dari localStorage
    let history = JSON.parse(localStorage.getItem('dataHistory')) || [];
    
    // Tambahkan data baru dengan timestamp
    const now = new Date();
    const tanggal = now.toLocaleDateString('id-ID', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
    });
    const waktu = now.toLocaleTimeString('id-ID');
    
    history.push({
        lokasi: lokasi,
        tanggal: tanggal,
        waktu: waktu,
        data: currentData,
        rataRata: rataRata
    });
    
    // Simpan kembali ke localStorage
    localStorage.setItem('dataHistory', JSON.stringify(history));
    
    // Buat workbook Excel
    const workbook = XLSX.utils.book_new();
    
    // Persiapkan data untuk sheet
    const sheetData = [['Lokasi', 'Tanggal', 'Waktu', 'Titik 1', 'Titik 2', 'Titik 3', 'Titik 4', 'Titik 5', 'Titik 6', 'Titik 7', 'Titik 8', 'Titik 9', 'Titik 10', 'Rata-rata']];
    
    // Tambahkan semua history ke sheet
    history.forEach(record => {
        const row = [record.lokasi, record.tanggal, record.waktu, ...record.data, record.rataRata];
        sheetData.push(row);
    });
    
    // Buat sheet dan tambahkan ke workbook
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    
    // Set width kolom
    worksheet['!cols'] = [
        { wch: 15 }, // Lokasi
        { wch: 12 }, // Tanggal
        { wch: 12 }, // Waktu
        { wch: 10 }, // Titik 1-10
        { wch: 10 },
        { wch: 10 },
        { wch: 10 },
        { wch: 10 },
        { wch: 10 },
        { wch: 10 },
        { wch: 10 },
        { wch: 10 },
        { wch: 10 },
        { wch: 12 }  // Rata-rata
    ];
    
    XLSX.utils.book_append_sheet(workbook, worksheet, "Data History");
    
    // Download file
    XLSX.writeFile(workbook, "data_history.xlsx");
    
    // Reset tabel ke kosong
    resetTabel();
    
    // Hapus isi lokasi
    lokasiInput.value = '';
    lokasiInput.focus();
    
    // Tampilkan notifikasi
    alert('Data berhasil diekspor! File: data_history.xlsx\nTabel sudah direset.');
}

// Fungsi untuk reset tabel
function resetTabel() {
    // Kosongkan semua cell
    for (let i = 1; i <= 10; i++) {
        document.getElementById('cell' + i).innerText = '-';
    }
    
    // Reset array
    dataArray = [];
    
    // Reset counter
    cellCounter = 1;
    
    // Sembunyikan result div
    resultDiv.style.display = 'none';
    averageResult.innerText = '';
    
    // Fokus ke input
    myInput.focus();
}

// Event listener untuk button export
document.getElementById('exportBtn').addEventListener('click', exportToExcel);

// Event listener untuk button reset history
document.getElementById('resetHistoryBtn').addEventListener('click', function() {
    if (confirm('Apakah Anda yakin ingin menghapus semua history? Tindakan ini tidak bisa dibatalkan!')) {
        localStorage.removeItem('dataHistory');
        alert('Semua history sudah dihapus. File Excel berikutnya akan dimulai dari kosong.');
    }
});