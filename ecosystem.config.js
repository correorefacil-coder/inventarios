module.exports = {
  apps: [
    {
      name: 'inventarios-backend',
      script: 'backend/api.py',
      interpreter: 'python3',
      instances: 1,
      autorestart: true,
      watch: false,
      max_memory_restart: '500M',
      env: {
        PORT: 8080,
        PYTHONUNBUFFERED: '1'
      }
    }
  ]
};
