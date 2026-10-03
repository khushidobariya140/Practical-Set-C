-- =====================================================================
-- queries_mysql.sql  |  Set C - Training Performance Analysis
-- SQL dialect : MySQL 8.0   |   Run AFTER setup_mysql.sql
-- Execution order: S2a -> S2b -> S2c -> S3a -> S3b
-- Scores rounded to 2 decimals. Pass = score >= 50.
-- MySQL Workbench shows each result in its own tab ("Result 1", "Result 2"...)
-- when you click Execute All.
-- =====================================================================

USE set_c_training;

-- S2a: Average score by department (lowest first)
SELECT  c.department,
        ROUND(AVG(a.score), 2) AS avg_score
FROM    assessments a
JOIN    courses c ON c.course_id = a.course_id
GROUP BY c.department
ORDER BY avg_score ASC, c.department ASC;

-- S2b: Courses whose average score is below 60
SELECT  c.course_id,
        c.course,
        c.department,
        ROUND(AVG(a.score), 2) AS avg_score
FROM    assessments a
JOIN    courses c ON c.course_id = a.course_id
GROUP BY c.course_id, c.course, c.department
HAVING  AVG(a.score) < 60
ORDER BY avg_score ASC, c.course_id ASC;

-- S2c: Top two batches by average score (ties broken alphabetically)
SELECT  batch,
        ROUND(AVG(score), 2) AS avg_score
FROM    assessments
GROUP BY batch
ORDER BY AVG(score) DESC, batch ASC
LIMIT 2;

-- S3a: LEFT JOIN courses -> assessments: number of assessment rows per course
SELECT  c.course_id,
        c.course,
        COUNT(a.assessment_id) AS assessment_rows
FROM    courses c
LEFT JOIN assessments a ON a.course_id = c.course_id
GROUP BY c.course_id, c.course
ORDER BY c.course_id;

-- S3b: Fact rows whose course_id has NO lookup match (expected: 0)
SELECT  COUNT(*) AS unmatched_keys
FROM    assessments a
LEFT JOIN courses c ON c.course_id = a.course_id
WHERE   c.course_id IS NULL;
