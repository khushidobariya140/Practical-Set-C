-- =====================================================================
-- setup_mysql.sql  |  Set C - Training Performance Analysis
-- SQL dialect : MySQL 8.0 (tested on 8.0.x; needs 8.0.16+ for CHECK)
-- Run order   : 1) setup_mysql.sql   2) queries_mysql.sql
-- Loads       : 4 course rows + 12 unique assessment rows
--               (the duplicate row for assessment_id 12 is excluded)
-- MySQL Workbench: File > Open SQL Script > click the lightning bolt
--                  (Execute All).
-- =====================================================================

CREATE DATABASE IF NOT EXISTS set_c_training;
USE set_c_training;

-- drop child table first because of the foreign key
DROP TABLE IF EXISTS assessments;
DROP TABLE IF EXISTS courses;

CREATE TABLE courses (
    course_id   VARCHAR(10)  NOT NULL,
    course      VARCHAR(50)  NOT NULL,
    department  VARCHAR(50)  NOT NULL,
    PRIMARY KEY (course_id)
) ENGINE = InnoDB;

CREATE TABLE assessments (
    assessment_id   INT           NOT NULL,
    month           VARCHAR(3)    NOT NULL,
    course_id       VARCHAR(10)   NOT NULL,
    batch           VARCHAR(20)   NOT NULL,
    score           DECIMAL(5,2)  NOT NULL,
    attendance_pct  DECIMAL(5,2)  NOT NULL,
    PRIMARY KEY (assessment_id),
    CONSTRAINT fk_assessments_course
        FOREIGN KEY (course_id) REFERENCES courses (course_id),
    CONSTRAINT chk_month  CHECK (month IN ('Jan','Feb','Mar')),
    CONSTRAINT chk_batch  CHECK (batch IN ('Morning','Evening','Weekend')),
    CONSTRAINT chk_score  CHECK (score BETWEEN 0 AND 100),
    CONSTRAINT chk_attend CHECK (attendance_pct BETWEEN 0 AND 100)
) ENGINE = InnoDB;

INSERT INTO courses (course_id, course, department) VALUES
    ('C1', 'Excel',   'Business'),
    ('C2', 'PowerBI', 'Business'),
    ('C3', 'SQL',     'Technology'),
    ('C4', 'Python',  'Technology');

INSERT INTO assessments (assessment_id, month, course_id, batch, score, attendance_pct) VALUES
    (1,  'Jan', 'C1', 'Morning', 72, 90),
    (2,  'Jan', 'C2', 'Evening', 45, 70),
    (3,  'Jan', 'C3', 'Morning', 65, 85),
    (4,  'Jan', 'C4', 'Weekend', 38, 60),
    (5,  'Feb', 'C1', 'Evening', 80, 95),
    (6,  'Feb', 'C2', 'Weekend', 55, 80),
    (7,  'Feb', 'C3', 'Morning', 48, 75),
    (8,  'Feb', 'C4', 'Evening', 68, 88),
    (9,  'Mar', 'C1', 'Weekend', 90, 98),
    (10, 'Mar', 'C2', 'Morning', 60, 82),
    (11, 'Mar', 'C3', 'Evening', 75, 92),
    (12, 'Mar', 'C4', 'Weekend', 42, 65);   -- duplicate 13th row intentionally NOT inserted

-- Row-count verification (expected: courses = 4, assessments = 12)
SELECT 'courses'     AS table_name, COUNT(*) AS row_count FROM courses
UNION ALL
SELECT 'assessments' AS table_name, COUNT(*) AS row_count FROM assessments;
