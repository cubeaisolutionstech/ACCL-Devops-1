from flask import Flask
from flask_cors import CORS
from config import Config
from extensions import db

from routes.mapping_routes import mapping_bp
from routes.bulk_assign_customers import bulk_bp
from routes.file_routes import file_bp
from routes.branch_region_routes import branch_region_bp
from routes.company_product_routes import company_product_bp
from routes.budget_routes import budget_bp
from routes.upload_tools import upload_tools_bp
from routes.sales_routes import sales_bp
from routes.os_processing_routes import os_bp

from routes.branch_routes import branch_bp
from routes.executive_routes import executive_bp

from routes.cumulative_routes import api_bp

from routes.auditor.auditor import auditor_bp
from routes.auditor.combined_data import combined_bp
from routes.auditor.data_routes import data_bp
from routes.auditor.ero_pw import ero_pw_bp
from routes.auditor.process_routes import process_bp
from routes.auditor.product import product_bp
from routes.auditor.region import region_bp
from routes.auditor.sales import sales1_bp
from routes.auditor.salesmonthwise import salesmonthwise_bp
from routes.auditor.ts_pw import ts_pw_bp
from routes.auditor.upload_routes import upload_bp

from routes.dashboard.main_routes import main_bp
from routes.routes import api1_bp

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    db.init_app(app)
    CORS(app, resources={r"/api/*": {"origins": ["http://localhost:3000"],
                                     "allow_headers": ["Content-Type"],
                                     "supports_credentials": True}
                                     })

    app.register_blueprint(mapping_bp, url_prefix="/api")
    app.register_blueprint(bulk_bp, url_prefix="/api")
    app.register_blueprint(file_bp, url_prefix="/api")
    app.register_blueprint(branch_region_bp, url_prefix="/api")
    app.register_blueprint(company_product_bp, url_prefix="/api")
    app.register_blueprint(budget_bp, url_prefix="/api")
    app.register_blueprint(upload_tools_bp, url_prefix="/api")
    app.register_blueprint(sales_bp, url_prefix="/api")
    app.register_blueprint(os_bp, url_prefix="/api")

    app.register_blueprint(branch_bp, url_prefix='/api/branch')
    app.register_blueprint(executive_bp, url_prefix='/api/executive')

    app.register_blueprint(api_bp,url_prefix="/api")

    app.register_blueprint(upload_bp, url_prefix='/api')
    app.register_blueprint(process_bp, url_prefix='/api')
    app.register_blueprint(data_bp, url_prefix='/api')
    app.register_blueprint(auditor_bp, url_prefix='/api')        
    app.register_blueprint(sales1_bp, url_prefix='/api')
    app.register_blueprint(region_bp, url_prefix='/api/region')
    app.register_blueprint(product_bp, url_prefix='/api/product')
    app.register_blueprint(ts_pw_bp, url_prefix='/api/ts-pw')
    app.register_blueprint(combined_bp, url_prefix='/api/combined')  
    app.register_blueprint(salesmonthwise_bp, url_prefix='/api')
    app.register_blueprint(ero_pw_bp, url_prefix='/api/ero-pw')

    app.register_blueprint(main_bp)
    app.register_blueprint(api1_bp, url_prefix='/api')
    
    
    return app

# db = SQLAlchemy()

if __name__ == "__main__":
    app = create_app()
    app.app_context().push()
    app.run(debug=True, port =5000)
