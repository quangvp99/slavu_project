-- drop database oauth2_sso;
-- create database oauth2_sso;
SET FOREIGN_KEY_CHECKS = 0;
drop table if exists role_permission;
drop table if exists user_role;
drop table if exists app_role;
drop table if exists app_user;
drop table if exists role_menu;
drop table if exists app_menu;
drop table if exists app_permission;
drop table if exists oauth_client_details;
SET FOREIGN_KEY_CHECKS = 1;

create table oauth_client_details (
  client_id VARCHAR(256) PRIMARY KEY,
  resource_ids VARCHAR(256),
  client_secret VARCHAR(256),
  scope VARCHAR(256),
  authorized_grant_types VARCHAR(256),
  web_server_redirect_uri VARCHAR(256),
  authorities VARCHAR(256),
  access_token_validity INTEGER,
  refresh_token_validity INTEGER,
  additional_information VARCHAR(4096),
  autoapprove VARCHAR(256),
  filter_show_on_client LONGTEXT DEFAULT NULL
)ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- client_secret: XY7kmzoNzl100
INSERT INTO oauth_client_details
   (client_id, client_secret, scope, authorized_grant_types,
   authorities, access_token_validity, refresh_token_validity)
VALUES
   ('Logistics', '$2a$10$QIE97SPFma2V3rmrL/rc2eJ8uxvRlkwIS0erxGMrJTpYidcGCd8im', 'read,write', 'password,refresh_token,client_credentials,authorization_code', 'ROLE_CLIENT,ROLE_TRUSTED_CLIENT', 900, 2592000);

-- client_secret: A31b4c24l3kj35d4AKJQ
INSERT INTO oauth_client_details
   (client_id, client_secret, scope, authorized_grant_types,
   authorities, access_token_validity, refresh_token_validity)
VALUES
   ('Dcommerce', '$2a$10$r7Ee6JriL4eZpV2Imdp8i.pWxVA9I5ZWPbc5A41NTP7er2M/B4QDy', 'read,write', 'password,refresh_token,client_credentials,authorization_code', 'ROLE_CLIENT,ROLE_TRUSTED_CLIENT', 900, 2592000);
   
-- client_secret: a15D6Ho2e3u4i7T8N9S0
INSERT INTO oauth_client_details
   (client_id, client_secret, scope, authorized_grant_types,
   authorities, access_token_validity, refresh_token_validity)
VALUES
   ('Hunter', '$2a$10$s9RtidRKBL41grm9i1QNoesFdvX2Wp0tFwhGvfpg1xi5TDW0CdFpK', 'read,write', 'password,refresh_token,client_credentials,authorization_code', 'ROLE_CLIENT,ROLE_TRUSTED_CLIENT', 900, 2592000);

-- client_secret: aoeuiDHTNS1234509876snthdIUEOA
INSERT INTO oauth_client_details
   (client_id, client_secret, scope, authorized_grant_types,
   authorities, access_token_validity, refresh_token_validity)
VALUES
   ('ShareWarehouse', '$2a$10$q8p9POkUZJ6ZACqWkjYtw.zuxrOaWXC4TooQgWQ1nI.RUSZSKw/KG', 'read,write', 'password,refresh_token,client_credentials,authorization_code', 'ROLE_CLIENT,ROLE_TRUSTED_CLIENT', 900, 2592000);
   
-- client_secret: 1qazXSW@3edcVFR$5tgbNHY^7ujm<KI*9ol.?:P)
INSERT INTO oauth_client_details
   (client_id, client_secret, scope, authorized_grant_types,
   authorities, access_token_validity, refresh_token_validity)
VALUES
   ('ShareVan', '$2a$10$8rhyG2nDUhWYTiCltToctO18Khn/OFcwvp5hQtjFOtVi0F8w6CzOm', 'read,write', 'password,refresh_token,client_credentials,authorization_code', 'ROLE_CLIENT,ROLE_TRUSTED_CLIENT', 900, 2592000);

-- client_secret: 2qazXSW@3edcVFR$5tgbNHY^7ujm<KI*9ol.?:P)
INSERT INTO oauth_client_details
   (client_id, client_secret, scope, authorized_grant_types,
   authorities, access_token_validity, refresh_token_validity)
VALUES
   ('Billing', '$2a$10$lU1kBdPdCrFKZ8TbkeEbqucsizzuFIs6nXjrYdOUAU5e0JLFNTqAu', 'read,write', 'password,refresh_token,client_credentials,authorization_code', 'ROLE_CLIENT,ROLE_TRUSTED_CLIENT', 900, 2592000);

CREATE TABLE app_user (
  id bigint NOT NULL AUTO_INCREMENT,
  first_name varchar(255) NOT NULL,
  last_name varchar(255) NOT NULL,
  password varchar(255) NOT NULL,
  username varchar(255) UNIQUE COLLATE utf8_unicode_ci NOT NULL,
  attempt_login_failed tinyint DEFAULT NULL,
  enabled bit(1) NOT NULL default false,
  must_change_password bit(1) NOT NULL default false,
  PRIMARY KEY (id)
)ENGINE=InnoDB AUTO_INCREMENT=1000 DEFAULT CHARSET=utf8;


-- USER
-- non-encrypted password: abc@123
-- DELETE FROM user_role where user_id < 100;
-- DELETE from app_user where id < 100;
INSERT INTO app_user (id, first_name, last_name, password, username, enabled)
            VALUES (1, 'GOD', 'GOD', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'GOD', 1);
-- INSERT INTO app_user (id, first_name, last_name, password, username, enabled)
--             VALUES (1, 'Merchant', 'Merchant', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'merchant', 1),
--                    (2, 'Admin', 'Admin', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'admin', 1),
--                    (3, 'StoreKeeper', 'StoreKeeper', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'storekeeper', 1),
--                    (4, 'GOD', 'GOD', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'GOD', 1),
--                    (5, 'phambien', 'phambien', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'phambien', 1),
--                    (6, 'hunter', 'hunter', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'hunter', 1),
--                    (7, 'du', 'du', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'du', 1),
--                    (8, 'fc_admin', 'fc_admin', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'fc_admin', 1),
--                    (9, 'fc_staff1', 'fc_staff1', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'fc_staff1', 1),
--                    (10, 'dc1_admin', 'dc1_admin', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'dc1_admin', 1),
--                    (11, 'dc1_staff1', 'dc1_staff1', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'dc1_staff1', 1),
--                    (12, 'dc2_admin', 'dc2_admin', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'dc2_admin', 1),
--                    (13, 'dc2_staff1', 'dc2_staff1', '$2a$10$dAPmXXaiF6/OirSNUDHnaOKC5YYjJdCcjK2eg7GcoOXkJ0RRvB0Qa', 'dc2_staff1', 1);


CREATE TABLE app_role (
  id bigint NOT NULL AUTO_INCREMENT,
  client_id VARCHAR(256) NOT NULL,
  description varchar(255) DEFAULT NULL,
  role_name varchar(255) UNIQUE NOT NULL,
  PRIMARY KEY (id)
)ENGINE=InnoDB AUTO_INCREMENT=1000 DEFAULT CHARSET=utf8;

-- DELETE FROM user_role where role_id < 50;
-- DELETE from app_role where id < 50;
-- INSERT INTO app_role (id, client_id, role_name, description)
--             VALUES (1, 'SharedWarehouse', 'MERCHANT', 'Merchant - Has no admin rights'),
--                    (2, 'SharedWarehouse', 'ADMIN', 'Admin - Has permission to perform admin tasks'),
--                    (3, 'SharedWarehouse', 'STORE_KEEPER', 'STORE KEEPER - Has permission to import/export statement'),
--                    (4, '', 'GOD', 'GOD OF WORLD'),
--                    (5, 'SharedWarehouse', 'STAFF', 'STAFF OF WORLD'),
--                    (6, 'Logistics', 'ADMIN_FC', ''),
--                    (7, 'Logistics', 'STAFF_FC', ''),
--                    (8, 'Logistics', 'ADMIN_DC', ''),
--                    (9, 'Logistics', 'STAFF_DC', '');


CREATE TABLE user_role (
  user_id bigint NOT NULL,
  role_id bigint NOT NULL,
  PRIMARY KEY (user_id, role_id),
  CONSTRAINT fk_userrole_user FOREIGN KEY (user_id) REFERENCES app_user (id),
  CONSTRAINT fk_userrole_role FOREIGN KEY (role_id) REFERENCES app_role (id)
);

-- INSERT INTO user_role(user_id, role_id) VALUES(1,1);
-- INSERT INTO user_role(user_id, role_id) VALUES(2,1);
-- INSERT INTO user_role(user_id, role_id) VALUES(2,2);
-- INSERT INTO user_role(user_id, role_id) VALUES(3,3);
-- INSERT INTO user_role(user_id, role_id) VALUES(4,4);
-- INSERT INTO user_role(user_id, role_id) VALUES(7,5);
-- 
-- INSERT INTO user_role(user_id, role_id)
--             VALUES(8,6),(9,7),(10,8),(11,9),(12,8),(13,9);



create table app_permission (
  id bigint NOT NULL AUTO_INCREMENT,
  client_id VARCHAR(256) NOT NULL,
  url varchar(255) NOT NULL,
  description longtext DEFAULT NULL,
  PRIMARY KEY (id)
)ENGINE=InnoDB AUTO_INCREMENT=1000 DEFAULT CHARSET=utf8;
            
            

CREATE TABLE role_permission (
  role_id bigint NOT NULL,
  permission_id bigint NOT NULL,
  PRIMARY KEY (role_id, permission_id),
  CONSTRAINT fk_rolepermission_role FOREIGN KEY (role_id) REFERENCES app_role (id),
  CONSTRAINT fk_rolepermission_permission FOREIGN KEY (permission_id) REFERENCES app_permission (id)
);

-- -- MERCHANT
-- INSERT INTO role_permission(role_id, permission_id) VALUES(1,5),(1,21),(1,22),(1,23);
-- 
-- INSERT INTO role_permission(role_id, permission_id)
--             VALUES(1,104),(1,105),(1,120),(1,121),(1,124),(1,125),(1,126),(1,127),(1,128),(1,129),(1,130),(1,131),(1,132),
--                   (1,133),(1,134),(1,135),(1,118),(1,119),(1,115),(1,110),(1,111);
-- -- ADMIN
-- INSERT INTO role_permission(role_id, permission_id)
--             VALUES (2, 1),(2, 5),(2, 15),(2, 17),(2, 19),(2, 62),(2, 20),(2,22),(2, 16),(2, 18), (2, 63), (2, 64), (2, 65),
--                    (2, 66), (2, 67), (2, 68), (2, 69), (2, 70), (2,71), (2, 72);
-- 
-- 
-- INSERT INTO role_permission(role_id, permission_id)
--             VALUES(2,100),(2,101),(2,102),(2,103),(2,104),(2,105),(2,106),(2,107),(2,108),(2,109),(2,110),(2,111),(2,112),
--                   (2,113),(2,114),(2,115),(2,116),(2,117),(2,118),(2,119),(2,120),(2,121),(2,122),(2,123),(2,124),(2,125),
--                   (2,126),(2,127),(2,128),(2,129),(2,130),(2,131),(2,132),(2,133),(2,134),(2,135);
--             
-- -- STORE_KEEPER
-- INSERT INTO role_permission(role_id, permission_id)
--             VALUES(3,1),(3,2),(3,5),(3,10),(3,14),(3,19),(3,20),(3,21),(3,22),(3,23),(3,24),(3,25),(3,26),(3,27),(3,28),
--                   (3,29),(3,30),(3,31),(3,32),(3,33),(3,38),(3,39);
-- -- GOD
-- INSERT INTO role_permission(role_id, permission_id)
--             VALUES(4,1),(4,2),(4,3),(4,4),(4,5),(4,6),(4,7),(4,8),(4,9),(4,10),(4,11),(4,12),(4,13),(4,14),(4,15),
--                   (4,16),(4,17),(4,18),(4,19),(4,20),(4,21),(4,22),(4,23),(4,24),(4,25),(4,26),(4,27),(4,28),(4,29),(4,30),
--                   (4,31),(4,32),(4,33),(4,34),(4,35),(4,36),(4,37),(4,38),(4,39),(4,40),(4,41),(4,42),(4,43),(4,44),(4,45),
--                   (4,46),(4,47),(4,48),(4,49),(4,50),(4,51),(4,52),(4,53),(4,54),(4,55),(4,56),(4,57),(4,58),(4,59),(4,60),
--                   (4,61),(4,62),(4, 63),(4, 64),(4, 65),(4, 66),(4, 67),(4, 68),(4, 69),(4, 70),(4, 72),(4, 73),(4, 74),(4, 75);
--                   
-- INSERT INTO role_permission(role_id, permission_id)
--             VALUES(4,100),(4,101),(4,102),(4,103),(4,104),(4,105),(4,106),(4,107),(4,108),(4,109),(4,110),(4,111),(4,112),
--                   (4,113),(4,114),(4,115),(4,116),(4,117),(4,118),(4,119),(4,120),(4,121),(4,122),(4,123),(4,124),(4,125),
--                   (4,126),(4,127),(4,128),(4,129),(4,130),(4,131),(4,132),(4,133),(4,134),(4,135);
--             
-- -- STAFF
-- INSERT INTO role_permission(role_id, permission_id)
--             VALUES(5,100),(5,101),(5,102),(5,103),(5,104),(5,105),(5,106),(5,107),(5,110),(5,111),(5,112),(5,113);
--             
-- -- ADMIN_FC
-- INSERT INTO role_permission(role_id, permission_id) VALUES (6, 29),(6, 31),(6, 30),(6, 28),(6, 27),(6, 41),(6, 43),(6, 42),(6, 44),(6, 45),(6, 33),(6, 32),(6, 26),(6, 57),(6, 25),(6, 24),(6, 58),(6, 22),(6, 21),(6, 46),(6, 48),(6, 47),(6, 49),(6, 35),(6, 36),(6, 34),(6, 40),(6, 37),(6, 39),(6, 38),(6, 50),(6, 56),(6, 55),(6, 52),(6, 51),(6, 61),(6, 54),(6, 53),(6, 62),(6, 20),(6, 16),(6, 72),(6, 73),(6, 74),(6, 75);
-- 
--                   
-- -- STAFF_FC
-- INSERT INTO role_permission(role_id, permission_id) VALUES (7, 29),(7, 31),(7, 30),(7, 28),(7, 27),(7, 41),(7, 43),(7, 42),(7, 44),(7, 45),(7, 33),(7, 32),(7, 26),(7, 57),(7, 25),(7, 24),(7, 58),(7, 46),(7, 47),(7, 35),(7, 36),(7, 34),(7, 40),(7, 37),(7, 39),(7, 38),(7, 50),(7, 56),(7, 55),(7, 52),(7, 51),(7, 61),(7, 54),(7, 53),(7, 62),(7, 20),(7, 16),(7, 72),(7, 73),(7, 74),(7, 75);
-- 
-- 
-- -- ADMIN_DC
-- INSERT INTO role_permission(role_id, permission_id) VALUES (8, 29),(8, 31),(8, 30),(8, 28),(8, 27),(8, 33),(8, 32),(8, 26),(8, 25),(8, 24),(8, 39),(8, 38),(8, 62),(8, 20),(8, 16),(8, 72),(8, 73),(8, 74),(8, 75);
-- 
--                   
-- -- STAFF_DC
-- INSERT INTO role_permission(role_id, permission_id) VALUES (9, 29),(9, 31),(9, 30),(9, 28),(9, 27),(9, 33),(9, 32),(9, 26),(9, 25),(9, 24),(9, 39),(9, 38),(9, 62),(9, 20),(9, 16),(9, 72),(9, 73),(9, 74),(9, 75);




create table app_menu(
  id bigint NOT NULL AUTO_INCREMENT,
  client_id VARCHAR(256) NOT NULL,
  app_type varchar(255) NOT NULL COMMENT 'Constants: MOBILE, WEB',
  code varchar(255) default null COMMENT 'Mapping với fragment của MOBILE',
  url varchar(255) default null COMMENT 'Mapping với page của WEB',
  parent_menu_id bigint default null,
  PRIMARY KEY (id),
  CONSTRAINT fk_menu_menu FOREIGN KEY (parent_menu_id) REFERENCES app_menu(id)
)ENGINE=InnoDB AUTO_INCREMENT=1000 DEFAULT CHARSET=utf8;

            
CREATE TABLE role_menu (
  role_id bigint NOT NULL,
  menu_id bigint NOT NULL,
  PRIMARY KEY (role_id, menu_id),
  CONSTRAINT fk_rolemenu_role FOREIGN KEY (role_id) REFERENCES app_role (id),
  CONSTRAINT fk_rolemenu_menu FOREIGN KEY (menu_id) REFERENCES app_menu (id)
);   

-- -- ADMIN
-- INSERT INTO role_menu(role_id, menu_id) VALUES (2, 50),(2, 51),(2, 52),(2,65);
-- INSERT INTO role_menu(role_id, menu_id) VALUES (2, 100), (2, 101), (2, 102), (2, 103), (2, 104), (2, 105), (2, 106), (2, 108), (2, 109), (2, 111);
-- 
-- -- ADMIN_FC
-- INSERT INTO role_menu(role_id, menu_id) VALUES (6, 1),(6, 2),(6, 3),(6, 4),(6, 5),(6, 6),(6, 7),(6, 8),(6, 9),(6, 10),(6, 11),(6, 12),(6, 13),(6, 14),(6, 15),(6, 16),(6, 17),(6, 18);
-- INSERT INTO role_menu(role_id, menu_id) VALUES (6, 50),(6, 51),(6, 53),(6, 54),(6, 55),(6, 56),(6, 57),(6, 58),(6, 59),(6, 60),(6, 61),(6, 62),(6, 63),(6, 64),(6, 66),(6,67);
-- 
-- -- STAFF_FC
-- INSERT INTO role_menu(role_id, menu_id) VALUES (7, 2),(7, 3),(7, 4),(7, 10),(7, 11),(7, 12),(7, 15),(7, 16);
-- INSERT INTO role_menu(role_id, menu_id) VALUES (7, 50);
-- 
-- -- ADMIN_DC
-- INSERT INTO role_menu(role_id, menu_id) VALUES (8, 2),(8, 3),(8, 4),(8, 9),(8, 10),(8, 12),(8, 13),(8, 14),(8, 16),(8, 17);
-- INSERT INTO role_menu(role_id, menu_id) VALUES (8, 50),(8, 51),(8, 54),(8, 59),(8, 60),(8, 61),(8, 62),(8, 63),(8, 64),(8, 66);
-- 
-- -- STAFF_DC
-- INSERT INTO role_menu(role_id, menu_id) VALUES (9, 2),(9, 3),(9, 4),(9, 10),(9, 16),(9, 17);
-- INSERT INTO role_menu(role_id, menu_id) VALUES (9, 50);




drop table if exists oauth_client_token;
create table oauth_client_token (
token_id VARCHAR(255),
token BLOB ,
authentication_id VARCHAR(255),
user_name VARCHAR(255),
client_id VARCHAR(255)
);

drop table if exists oauth_access_token;
create table oauth_access_token (
token_id VARCHAR(255),
token BLOB,
authentication_id VARCHAR(255) PRIMARY KEY,
user_name VARCHAR(255),
client_id VARCHAR(255),
authentication BLOB,
refresh_token VARCHAR(255),
UNIQUE KEY unique_oauth_access_token(user_name, client_id)
);

create index oauth_access_token_id on oauth_access_token(token_id);
create index oauth_refresh_token_id on oauth_access_token(token_id);

drop table if exists oauth_refresh_token;
create table oauth_refresh_token (
token_id VARCHAR(255),
token BLOB,
authentication BLOB
);
drop table if exists oauth_code;
create table oauth_code (
code VARCHAR(255), authentication BLOB
);

drop table if exists oauth_approvals;
create table oauth_approvals (
userId VARCHAR(255),
clientId VARCHAR(255),
scope VARCHAR(255),
status VARCHAR(10),
expiresAt TIMESTAMP,
lastModifiedAt TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
drop table if exists ClientDetails;
create table ClientDetails (
appId VARCHAR(255) PRIMARY KEY,
resourceIds VARCHAR(255),
appSecret VARCHAR(255),
scope VARCHAR(255),
grantTypes VARCHAR(255),
redirectUrl VARCHAR(255),
authorities VARCHAR(255),
access_token_validity INTEGER,
refresh_token_validity INTEGER,
additionalInformation VARCHAR(4096),
autoApproveScopes VARCHAR(255)
);

DROP TABLE IF EXISTS login_log;

CREATE TABLE login_log
(
   id            BIGINT(20) NOT NULL AUTO_INCREMENT,
   username      VARCHAR(255) NOT NULL,
   log_date      DATETIME NOT NULL,
   `month`       TINYINT NOT NULL,
   ip_address    VARCHAR(50) DEFAULT NULL,
   status        ENUM('FAILED', 'SUCCESS') NOT NULL,
   PRIMARY KEY(id, log_date, `month`)
)
ENGINE = InnoDB
AUTO_INCREMENT = 1
DEFAULT CHARSET = utf8
PARTITION BY RANGE (YEAR(log_date))
   SUBPARTITION BY KEY (`month`)
      SUBPARTITIONS 12
   (PARTITION p2020 VALUES LESS THAN (2021),
    PARTITION p2021 VALUES LESS THAN (2022),
    PARTITION p2022 VALUES LESS THAN (2023),
    PARTITION p2023 VALUES LESS THAN (2024),
    PARTITION p2024 VALUES LESS THAN (2025),
    PARTITION p2025 VALUES LESS THAN (2026));

delete from login_log;
explain partitions select * from login_log partition(p2020) where month in(9, 10) order by log_date desc;

DROP TABLE IF EXISTS count_action_log;
CREATE TABLE count_action_log
(
   id            BIGINT(20) NOT NULL AUTO_INCREMENT,
   username      VARCHAR(255) NOT NULL,
   log_date      DATETIME NOT NULL,
   `month`       TINYINT NOT NULL,
   total_action  BIGINT(20) NOT NULL,
   PRIMARY KEY(id, log_date, `month`),
   UNIQUE KEY (username, log_date, `month`)
)
ENGINE = InnoDB
AUTO_INCREMENT = 1
DEFAULT CHARSET = utf8
PARTITION BY RANGE (YEAR(log_date))
   SUBPARTITION BY KEY (`month`)
      SUBPARTITIONS 12
   (PARTITION p2020 VALUES LESS THAN (2021),
    PARTITION p2021 VALUES LESS THAN (2022),
    PARTITION p2022 VALUES LESS THAN (2023),
    PARTITION p2023 VALUES LESS THAN (2024),
    PARTITION p2024 VALUES LESS THAN (2025),
    PARTITION p2025 VALUES LESS THAN (2026));
    
select * from count_action_log order by total_action desc;
 SELECT NVL(SUM(cal.total_action), 0) 
		 FROM count_action_log cal 
		 WHERE cal.username LIKE 'mstoreadmin'
			AND DATE(SYSDATE()) <= cal.log_date 
			AND cal.log_date < DATE(DATE_ADD(SYSDATE(), INTERVAL 1 DAY)) ;

delete from oauth_refresh_token;
delete from oauth_access_token;
commit;
                   
