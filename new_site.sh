#!/bin/bash

# Create the folder
read -p 'Folder name: ' foldername
mkdir $foldername
echo Created $foldername in /etc/docker-sites

# Create a new wordpress.yml file for the site
ports_in_use=$(ss -t -n -l | awk '{print $4}')
echo The following ports should not be used $ports_in_use
read -p "Port: " port
read -p "Table prefix: " prefix

# Send the wordpress.yml to the folder
cat <<EOF >$foldername/wordpress.yml
version: '3.1'
services:

  wordpress:
    image: wordpress:latest
    restart: always
    volumes:
      - ./wp-content:/var/www/html
    ports:
      - $port:80
    environment:
      WORDPRESS_DB_HOST: db
      WORDPRESS_DB_USER: wp_user
      WORDPRESS_DB_PASSWORD: CHANGEME
      WORDPRESS_DB_NAME: wp
      WORDPRESS_TABLE_PREFIX: $prefix

networks:
  default:
    external:
      name: dockersites_default

volumes:
  db_data:
EOF
