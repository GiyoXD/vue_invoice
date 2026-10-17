#!/bin/bash
set -e

# When using docker-compose, the host mounts an empty volume to /app/storage
# This script seeds the volume with the default configuration files (like mapping_config.json)
# if they don't already exist in the mounted volume.

if [ -d "/app/default_storage" ]; then
    echo "Updating system mapper configurations in volume..."
    mkdir -p /app/storage/blueprints/mapper
    mkdir -p /app/storage/db
    
    # Always update system mapping configurations (master_config.json) to the latest version
    if [ -f "/app/default_storage/blueprints/mapper/master_config.json" ]; then
        cp /app/default_storage/blueprints/mapper/master_config.json /app/storage/blueprints/mapper/
    fi

    # Seed global map (mapping_config.json) only if it doesn't already exist in the volume
    if [ ! -f "/app/storage/blueprints/mapper/mapping_config.json" ] && [ -f "/app/default_storage/blueprints/mapper/mapping_config.json" ]; then
        cp /app/default_storage/blueprints/mapper/mapping_config.json /app/storage/blueprints/mapper/
    fi

    # Seed all other database folders and files (like registry.db) only if they do not exist
    echo "Seeding other database files if missing..."
    cp -r -n /app/default_storage/* /app/storage/
fi

# Execute the main command (uvicorn)
exec "$@"
