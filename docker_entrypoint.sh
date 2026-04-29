#!/bin/bash
set -e

# When using docker-compose, the host mounts an empty volume to /app/database
# This script seeds the volume with the default configuration files (like mapping_config.json)
# if they don't already exist in the mounted volume.

if [ -d "/app/default_database" ]; then
    echo "Seeding default database configurations into volume..."
    # Copy recursively, without overwriting existing files (-n)
    cp -r -n /app/default_database/* /app/database/
fi

# Execute the main command (uvicorn)
exec "$@"
