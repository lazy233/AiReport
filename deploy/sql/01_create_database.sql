-- =============================================================================
-- AiReport：手动创建 PostgreSQL 库与账号（在已有 PostgreSQL 实例上执行）
-- =============================================================================
-- 说明：
-- 1. 以超级用户（如 postgres）在 psql 或客户端中执行；请按需修改用户名、库名、密码。
-- 2. 业务表（parsed_presentations、generation_histories、students 等）由应用首次
--    启动时通过 SQLAlchemy Base.metadata.create_all 自动创建，无需手工跑 DDL。
-- 3. 若密码含 @ : / 等特殊字符，写入 DATABASE_URL 时需做 URL 编码（或使用纯字母数字密码）。
-- =============================================================================

-- 创建登录角色（密码请改为强密码）
CREATE USER pptreport WITH PASSWORD '请改为强密码';

-- 创建数据库并指定属主
CREATE DATABASE ppt_report_platform OWNER pptreport;

-- 可选：限制该用户只能连接到自己的库（按安全策略选用）
-- REVOKE CONNECT ON DATABASE postgres FROM pptreport;
