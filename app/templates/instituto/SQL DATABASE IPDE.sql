CREATE DATABASE ipde; 
USE ipde;

CREATE TABLE usuarios (
    id BIGINT AUTO_INCREMENT PRIMARY KEY,
    usuario VARCHAR(15) NOT NULL,
    password VARCHAR(255) NOT NULL
);

INSERT INTO usuarios (usuario, password)
VALUES ('admin', '12345');

UPDATE USUARIOS
SET PASSWORD = '12345678' 
WHERE ID = 1;

SELECT * FROM usuarios;
TRUNCATE TABLE usage_events;

------------------------------------------------------
-- segundo cambio por integración de BCRYPT en back --
------------------------------------------------------

ALTER TABLE usuarios
MODIFY password VARCHAR(255) NOT NULL;

UPDATE usuarios 
SET PASSWORD = '$2a$12$eh.IQT6XJSYFCQUz9oG2QuUZ3ywHKSojcwVRPQg6QkjKh2mY6oh2C'
WHERE ID = 1;

------------------------------------------------------
----------- tercer cambio para el nivel 8 ------------
------------------------------------------------------

-- token_hash → NUNCA guardamos el token real
-- revoked → control de sesión
-- Escalable a múltiples dispositivos

SELECT * FROM refresh_tokens;

CREATE TABLE refresh_tokens (
    id BIGINT AUTO_INCREMENT PRIMARY KEY,
    usuario VARCHAR(50) NOT NULL,
    token_hash VARCHAR(255) NOT NULL,
    revoked BOOLEAN DEFAULT FALSE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

------------------------------------------------------
----------- cuarto cambio para el nivel 9 ------------
------------------------------------------------------

ALTER TABLE refresh_tokens
ADD COLUMN device_info VARCHAR(255),
ADD COLUMN ip_address VARCHAR(45),
ADD COLUMN expires_at TIMESTAMP;

------------------------------------------------------
---------- quinto cambio para el nivel 9.1 -----------
------------------------------------------------------

SELECT * FROM usage_events;

TRUNCATE TABLE usage_events;


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

------------------------------------------------------
----------- quinto cambio para el nivel 10 -----------
------------------------------------------------------

ALTER TABLE usuarios
ADD COLUMN role VARCHAR(10) NOT NULL DEFAULT 'USER';

-- Tu usuario admin actual (id=1) debe ser ADMIN
UPDATE usuarios
SET role = 'ADMIN'
WHERE id = 1;













SELECT * FROM usuarios;