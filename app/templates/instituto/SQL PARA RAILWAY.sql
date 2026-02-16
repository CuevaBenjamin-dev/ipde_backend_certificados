SHOW DATABASES;
CREATE DATABASE ipde;
USE ipde;


CREATE TABLE usuarios (
    id BIGINT AUTO_INCREMENT PRIMARY KEY,
    usuario VARCHAR(15) NOT NULL,
    password VARCHAR(255) NOT NULL
);

ALTER TABLE usuarios
ADD COLUMN role VARCHAR(10) NOT NULL DEFAULT 'USER';

INSERT INTO usuarios (usuario, password, role)
VALUES ('admin', '$2a$12$eh.IQT6XJSYFCQUz9oG2QuUZ3ywHKSojcwVRPQg6QkjKh2mY6oh2C', 'ADMIN');

SELECT * FROM usuarios;
SELECT * FROM usage_events;


CREATE TABLE refresh_tokens (
    id BIGINT AUTO_INCREMENT PRIMARY KEY,
    usuario VARCHAR(50) NOT NULL,
    token_hash VARCHAR(255) NOT NULL,
    revoked BOOLEAN DEFAULT FALSE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

ALTER TABLE refresh_tokens
ADD COLUMN device_info VARCHAR(255),
ADD COLUMN ip_address VARCHAR(45),
ADD COLUMN expires_at TIMESTAMP;


CREATE TABLE usage_events (
    id BIGINT AUTO_INCREMENT PRIMARY KEY,

    usuario VARCHAR(50) NOT NULL,
    evento VARCHAR(50) NOT NULL,

    items_count INT NOT NULL,

    origen VARCHAR(20) DEFAULT 'WEB',

    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

    INDEX idx_usuario_fecha (usuario, created_at),
    INDEX idx_evento_fecha (evento, created_at)
);










