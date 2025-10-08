from flask import Flask, redirect
from routes.customers import customers_bp
from routes.products import products_bp
from routes.orders import orders_bp

app = Flask(__name__)

# Register blueprints
app.register_blueprint(customers_bp)
app.register_blueprint(products_bp)
app.register_blueprint(orders_bp)

@app.route('/')
def home():
    return redirect('/customers/add')

if __name__ == '__main__':
    app.run(debug=True)
