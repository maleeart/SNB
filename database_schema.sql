-- ==============================================================================
-- ระบบลงทะเบียนเข้าใช้งานอาคารและกิจกรรมกีฬา (Building & Sports Club Access System)
-- อาคารอเนกประสงค์ ต.0025 การไฟฟ้าฝ่ายผลิตแห่งประเทศไทย สำนักงานไทรน้อย
-- Database Schema for PostgreSQL / Supabase / MySQL
-- ==============================================================================

-- 1. ตารางข้อมูลชมรมและกิจกรรม (Clubs & Activities Table)
CREATE TABLE IF NOT EXISTS clubs (
    id VARCHAR(50) PRIMARY KEY,
    name_th VARCHAR(100) NOT NULL,
    name_en VARCHAR(100) NOT NULL,
    category VARCHAR(50) DEFAULT 'sports',
    icon VARCHAR(10) DEFAULT '🏢',
    theme_color VARCHAR(20) DEFAULT '#3b82f6',
    zone_description VARCHAR(255),
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- ข้อมูลชมรมเริ่มต้น (Default Clubs)
INSERT INTO clubs (id, name_th, name_en, category, icon, theme_color, zone_description)
VALUES
    ('basketball', 'ชมรมบาสเก็ตบอล', 'Basketball Club', 'sports', '🏀', '#f97316', 'คอร์ตบาสเก็ตบอลในร่ม'),
    ('golf', 'ชมรมกอล์ฟ', 'Golf Club', 'sports', '⛳', '#10b981', 'สนามไดรฟ์กอล์ฟและกรีนพัตต์'),
    ('archery', 'ชมรมยิงธนู', 'Archery Club', 'sports', '🎯', '#ef4444', 'ลู่ยิงธนูมาตรฐาน'),
    ('badminton', 'ชมรมแบดมินตัน', 'Badminton Club', 'sports', '🏸', '#06b6d4', 'คอร์ตแบดมินตัน 1-3'),
    ('heights_training', 'อบรมหลักสูตรการทำงานบนที่สูง', 'Working at Heights Training', 'training', '🦺', '#eab308', 'ศูนย์ฝึกอบรมการทำงานบนที่สูง / เสาจำลอง')
ON CONFLICT (id) DO NOTHING;

-- 2. ตารางผู้ใช้งาน / สมาชิก (Users & Members Table)
CREATE TABLE IF NOT EXISTS members (
    id VARCHAR(50) PRIMARY KEY,
    emp_id VARCHAR(50),
    full_name VARCHAR(150) NOT NULL,
    department VARCHAR(100) NOT NULL,
    default_club_id VARCHAR(50) REFERENCES clubs(id) ON DELETE SET NULL,
    phone_number VARCHAR(30),
    email VARCHAR(100),
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- 3. ตารางประวัติการลงทะเบียนเข้าใช้งานอาคาร (Building Access Logs Table)
CREATE TABLE IF NOT EXISTS access_logs (
    id VARCHAR(100) PRIMARY KEY,
    member_id VARCHAR(50) REFERENCES members(id) ON DELETE SET NULL,
    user_name VARCHAR(150) NOT NULL,
    department VARCHAR(100) NOT NULL,
    emp_id VARCHAR(50),
    club_id VARCHAR(50) NOT NULL REFERENCES clubs(id) ON DELETE CASCADE,
    club_name VARCHAR(100) NOT NULL,
    check_in_time TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    check_out_time TIMESTAMP WITH TIME ZONE,
    duration_minutes INTEGER,
    status VARCHAR(20) NOT NULL DEFAULT 'active', -- 'active', 'completed', 'cancelled'
    purpose VARCHAR(100) DEFAULT 'general',       -- 'practice', 'match', 'training', 'meeting'
    note TEXT,
    ip_address VARCHAR(50),
    user_agent TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Indexes เพื่อเพิ่มประสิทธิภาพการค้นหาและทำรายงานสรุป
CREATE INDEX IF NOT EXISTS idx_logs_club_id ON access_logs(club_id);
CREATE INDEX IF NOT EXISTS idx_logs_check_in ON access_logs(check_in_time);
CREATE INDEX IF NOT EXISTS idx_logs_status ON access_logs(status);
CREATE INDEX IF NOT EXISTS idx_logs_dept ON access_logs(department);
CREATE INDEX IF NOT EXISTS idx_logs_name ON access_logs(user_name);

-- 4. View สำหรับดูสถานะผู้ใช้งานปัจจุบันที่กำลังอยู่ในอาคาร (Live Active Occupants)
CREATE OR REPLACE VIEW view_active_occupants AS
SELECT 
    l.id AS log_id,
    l.user_name,
    l.department,
    l.club_id,
    c.name_th AS club_name,
    c.icon AS club_icon,
    c.theme_color,
    l.check_in_time,
    ROUND(EXTRACT(EPOCH FROM (CURRENT_TIMESTAMP - l.check_in_time)) / 60) AS minutes_in_facility,
    l.note
FROM access_logs l
JOIN clubs c ON l.club_id = c.id
WHERE l.status = 'active'
ORDER BY l.check_in_time DESC;

-- 5. View สำหรับสรุปสถิติตามชมรม (Club Statistics Summary)
CREATE OR REPLACE VIEW view_club_stats AS
SELECT 
    c.id AS club_id,
    c.name_th AS club_name,
    c.icon AS club_icon,
    COUNT(l.id) AS total_visits,
    COUNT(CASE WHEN l.status = 'active' THEN 1 END) AS currently_active,
    COALESCE(AVG(l.duration_minutes) FILTER (WHERE l.status = 'completed'), 0)::INTEGER AS avg_duration_minutes,
    COUNT(DISTINCT l.user_name) AS unique_visitors
FROM clubs c
LEFT JOIN access_logs l ON c.id = l.club_id
GROUP BY c.id, c.name_th, c.icon
ORDER BY total_visits DESC;
