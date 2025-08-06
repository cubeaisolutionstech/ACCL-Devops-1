from flask import Flask
from flask_cors import CORS
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG)

def create_app():
    app = Flask(__name__)
    CORS(app, origins=["http://localhost:3000"], supports_credentials=True)

    # Configuration
    app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload size
    app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}
    app.config['BRANCH_EXCLUDE_TERMS'] = [
        'CHN Total', 'ERD SALES', 'North Total', 'WEST SALES', 'GROUP COMPANIES'
    ]

    # Register blueprints
    from routes.dashboard.main_routes import main_bp
    app.register_blueprint(main_bp)

    return app


if __name__ == '__main__':
    app = create_app()
    print("🚀 Starting Flask Backend...")
    print("📊 Excel Dashboard API Server")
    print("🔗 Backend running on: http://localhost:5003")
    print("🧪 Test endpoint: http://localhost:5003/test")
    app.run(debug=True, host='0.0.0.0', port=5003) 