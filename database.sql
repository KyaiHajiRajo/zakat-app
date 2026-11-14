-- ========================================
-- Aplikasi Pendataan Zakat Fitrah
-- Database Setup Script
-- ========================================

-- Buat database (jalankan terpisah jika belum ada)
-- CREATE DATABASE zakat_fitrah CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
-- USE zakat_fitrah;

-- ========================================
-- Tabel Users (Panitia dan Admin)
-- ========================================
CREATE TABLE users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    nomor_wa VARCHAR(15) UNIQUE NOT NULL,
    password VARCHAR(255) NOT NULL,
    role ENUM('admin','panitia') DEFAULT 'panitia',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

-- ========================================
-- Tabel RT (Rukun Tetangga)
-- ========================================
CREATE TABLE rt (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nomor_rt VARCHAR(10) NOT NULL,          -- contoh: RT 11
    ketua_rt VARCHAR(100) NOT NULL,         -- nama ketua RT
    keterangan VARCHAR(100) NULL,           -- tambahan catatan opsional
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

-- ========================================
-- Tabel Muzakki (Pembayar Zakat)
-- ========================================
CREATE TABLE muzakki (
    id INT AUTO_INCREMENT PRIMARY KEY,
    rt_id INT NOT NULL,                         -- RT asal muzakki
    nama VARCHAR(100) NOT NULL,                 -- nama muzakki
    jumlah_jiwa INT NOT NULL,                   -- berapa jiwa yang ditanggung
    jenis_zakat ENUM('beras','uang') NOT NULL,  -- jenis zakat
    jumlah_beras_kg DECIMAL(10,2) DEFAULT NULL, -- total kewajiban beras (kg)
    jumlah_uang DECIMAL(10,2) DEFAULT NULL,     -- total kewajiban uang (Rp)
    jumlah_bayar DECIMAL(10,2) DEFAULT NULL,    -- jumlah yang dibayar
    kembalian DECIMAL(10,2) DEFAULT 0,          -- otomatis dihitung
    catatan VARCHAR(255) NULL,                  -- keterangan opsional
    user_id INT NOT NULL,                       -- panitia pencatat
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (rt_id) REFERENCES rt(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

-- ========================================
-- Tabel Mustahik (Penerima Zakat)
-- ========================================
CREATE TABLE mustahik (
    id INT AUTO_INCREMENT PRIMARY KEY,
    rt_id INT DEFAULT NULL,
    nama VARCHAR(100) NOT NULL,
    kategori ENUM('fakir','miskin','amil','mualaf','fisabilillah','gharim','ibnusabil') NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (rt_id) REFERENCES rt(id) ON DELETE SET NULL
);

-- ========================================
-- Tabel Distribusi Zakat
-- ========================================
CREATE TABLE distribusi_zakat (
    id INT AUTO_INCREMENT PRIMARY KEY,
    mustahik_id INT NOT NULL,
    jenis_zakat ENUM('beras','uang') NOT NULL,
    jumlah DECIMAL(10,2) NOT NULL,
    status ENUM('pending','disalurkan','diterima','batal') DEFAULT 'pending',
    user_id INT NOT NULL,  -- panitia yang menyalurkan
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (mustahik_id) REFERENCES mustahik(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

-- ========================================
-- Tabel Infak (dari kembalian zakat)
-- ========================================
CREATE TABLE infak (
    id INT AUTO_INCREMENT PRIMARY KEY,
    muzakki_id INT NOT NULL,
    jumlah DECIMAL(10,2) NOT NULL,
    keterangan VARCHAR(255) NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (muzakki_id) REFERENCES muzakki(id) ON DELETE CASCADE
);

-- ========================================
-- Index untuk Performance
-- ========================================
CREATE INDEX idx_muzakki_rt ON muzakki(rt_id);
CREATE INDEX idx_muzakki_user ON muzakki(user_id);
CREATE INDEX idx_muzakki_jenis ON muzakki(jenis_zakat);
CREATE INDEX idx_infak_muzakki ON infak(muzakki_id);
CREATE INDEX idx_distribusi_mustahik ON distribusi_zakat(mustahik_id);
CREATE INDEX idx_users_nomor_wa ON users(nomor_wa);

-- ========================================
-- Sample Data untuk Testing
-- ========================================

-- Sample RT Data
INSERT INTO rt (nomor_rt, ketua_rt, keterangan) VALUES
('RT 01', 'Bapak Ahmad Santoso', 'Wilayah Utara - Jl. Mawar'),
('RT 02', 'Bapak Budi Raharja', 'Wilayah Selatan - Jl. Melati'),
('RT 03', 'Bapak Candra Wijaya', 'Wilayah Timur - Jl. Anggrek'),
('RT 04', 'Ibu Dewi Sartika', 'Wilayah Barat - Jl. Kenanga'),
('RT 05', 'Bapak Eko Prasetyo', 'Wilayah Tengah - Jl. Dahlia');

-- Sample Users (Password semua: admin123)
-- Hash untuk 'admin123' = $2a$10$92IXUNpkjO0rOQ5byMi.Ye4oKoEa3Ro9llC/.og/at2.uheWG/igi
INSERT INTO users (name, nomor_wa, password, role) VALUES
('Admin Zakat', '081935679330', '$2a$10$92IXUNpkjO0rOQ5byMi.Ye4oKoEa3Ro9llC/.og/at2.uheWG/igi', 'admin'),
('Panitia RT 01', '08234567890', '$2a$10$92IXUNpkjO0rOQ5byMi.Ye4oKoEa3Ro9llC/.og/at2.uheWG/igi', 'panitia'),
('Panitia RT 02', '08345678901', '$2a$10$92IXUNpkjO0rOQ5byMi.Ye4oKoEa3Ro9llC/.og/at2.uheWG/igi', 'panitia'),
('Panitia RT 03', '08456789012', '$2a$10$92IXUNpkjO0rOQ5byMi.Ye4oKoEa3Ro9llC/.og/at2.uheWG/igi', 'panitia');

-- Sample Muzakki Data
INSERT INTO muzakki (rt_id, nama, jumlah_jiwa, jenis_zakat, jumlah_beras_kg, jumlah_uang, jumlah_bayar, kembalian, catatan, user_id) VALUES
(1, 'Keluarga Ahmad', 4, 'beras', 10.0, NULL, 150000, 0, 'Bayar cash', 2),
(1, 'Keluarga Siti', 3, 'uang', NULL, 135000, 150000, 15000, 'Ada kembalian', 2),
(2, 'Keluarga Budi', 5, 'beras', 12.5, NULL, 200000, 0, 'Lengkap', 3),
(2, 'Keluarga Ani', 2, 'uang', NULL, 90000, 100000, 10000, 'Kembalian 10rb', 3),
(3, 'Keluarga Candra', 6, 'beras', 15.0, NULL, 300000, 0, 'Keluarga besar', 4);

-- Sample Mustahik Data
INSERT INTO mustahik (rt_id, nama, kategori) VALUES
(1, 'Pak Umar (Fakir)', 'fakir'),
(1, 'Bu Fatimah (Miskin)', 'miskin'),
(2, 'Pak Ali (Fisabilillah)', 'fisabilillah'),
(3, 'Bu Khadijah (Amil)', 'amil'),
(NULL, 'Yayasan Yatim', 'fakir');

-- Sample Infak dari kembalian
INSERT INTO infak (muzakki_id, jumlah, keterangan) VALUES
(2, 15000, 'Kembalian dari zakat fitrah'),
(4, 10000, 'Kembalian dari zakat fitrah');

-- ========================================
-- Views untuk Laporan (Opsional)
-- ========================================

-- View Laporan per RT
CREATE VIEW view_laporan_rt AS
SELECT 
    r.id as rt_id,
    r.nomor_rt,
    r.ketua_rt,
    COUNT(m.id) as total_muzakki,
    SUM(CASE WHEN m.jenis_zakat = 'beras' THEN m.jumlah_beras_kg ELSE 0 END) as total_beras_kg,
    SUM(CASE WHEN m.jenis_zakat = 'uang' THEN m.jumlah_uang ELSE 0 END) as total_uang_kewajiban,
    SUM(m.jumlah_bayar) as total_terkumpul,
    SUM(m.kembalian) as total_kembalian
FROM rt r
LEFT JOIN muzakki m ON r.id = m.rt_id
GROUP BY r.id, r.nomor_rt, r.ketua_rt
ORDER BY r.nomor_rt;

-- View Total Infak
CREATE VIEW view_total_infak AS
SELECT 
    COUNT(*) as total_transaksi,
    SUM(jumlah) as total_infak,
    AVG(jumlah) as rata_rata_infak
FROM infak;

-- ========================================
-- Stored Procedures (Opsional)
-- ========================================

DELIMITER //

-- Procedure untuk menghitung statistik zakat
CREATE PROCEDURE GetZakatStats()
BEGIN
    SELECT 
        'Total Muzakki' as kategori,
        COUNT(*) as jumlah,
        NULL as nilai
    FROM muzakki
    
    UNION ALL
    
    SELECT 
        'Total Zakat Beras' as kategori,
        NULL as jumlah,
        CONCAT(SUM(jumlah_beras_kg), ' kg') as nilai
    FROM muzakki 
    WHERE jenis_zakat = 'beras'
    
    UNION ALL
    
    SELECT 
        'Total Zakat Uang' as kategori,
        NULL as jumlah,
        CONCAT('Rp ', FORMAT(SUM(jumlah_bayar), 0)) as nilai
    FROM muzakki 
    WHERE jenis_zakat = 'uang'
    
    UNION ALL
    
    SELECT 
        'Total Infak' as kategori,
        NULL as jumlah,
        CONCAT('Rp ', FORMAT(SUM(jumlah), 0)) as nilai
    FROM infak;
END //

DELIMITER ;

-- ========================================
-- Triggers untuk Audit Log (Opsional)
-- ========================================

-- Log perubahan data muzakki
CREATE TABLE audit_muzakki (
    id INT AUTO_INCREMENT PRIMARY KEY,
    muzakki_id INT,
    action ENUM('INSERT', 'UPDATE', 'DELETE'),
    old_data JSON,
    new_data JSON,
    user_id INT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

DELIMITER //

CREATE TRIGGER muzakki_audit_insert AFTER INSERT ON muzakki
FOR EACH ROW
BEGIN
    INSERT INTO audit_muzakki (muzakki_id, action, new_data, user_id)
    VALUES (NEW.id, 'INSERT', JSON_OBJECT(
        'nama', NEW.nama,
        'jumlah_jiwa', NEW.jumlah_jiwa,
        'jenis_zakat', NEW.jenis_zakat,
        'jumlah_bayar', NEW.jumlah_bayar
    ), NEW.user_id);
END //

CREATE TRIGGER muzakki_audit_update AFTER UPDATE ON muzakki
FOR EACH ROW
BEGIN
    INSERT INTO audit_muzakki (muzakki_id, action, old_data, new_data, user_id)
    VALUES (NEW.id, 'UPDATE', 
        JSON_OBJECT(
            'nama', OLD.nama,
            'jumlah_jiwa', OLD.jumlah_jiwa,
            'jenis_zakat', OLD.jenis_zakat,
            'jumlah_bayar', OLD.jumlah_bayar
        ),
        JSON_OBJECT(
            'nama', NEW.nama,
            'jumlah_jiwa', NEW.jumlah_jiwa,
            'jenis_zakat', NEW.jenis_zakat,
            'jumlah_bayar', NEW.jumlah_bayar
        ), 
        NEW.user_id
    );
END //

DELIMITER ;

-- ========================================
-- Query Contoh untuk Testing
-- ========================================

-- Lihat semua muzakki dengan info RT
-- SELECT m.*, r.nomor_rt, r.ketua_rt, u.name as pencatat 
-- FROM muzakki m 
-- LEFT JOIN rt r ON m.rt_id = r.id 
-- LEFT JOIN users u ON m.user_id = u.id;

-- Laporan per RT
-- SELECT * FROM view_laporan_rt;

-- Total infak
-- SELECT * FROM view_total_infak;

-- Muzakki dengan kembalian
-- SELECT nama, kembalian FROM muzakki WHERE kembalian > 0;

-- ========================================
-- Backup Command (untuk dokumentasi)
-- ========================================
-- mysqldump -u root -p zakat_fitrah > backup_zakat_fitrah.sql

-- ========================================
-- Restore Command (untuk dokumentasi)  
-- ========================================
-- mysql -u root -p zakat_fitrah < backup_zakat_fitrah.sql