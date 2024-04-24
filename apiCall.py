from flask import Flask, jsonify
from prepareDeck import prepareDeck
app = Flask(__name__)

# Define a route for the GET request
@app.route('/', methods=['GET'])
def hello():
    prepareDeck()
    return jsonify({'message': 'Hello, World!'})

if __name__ == '__main__':
    app.run(debug=False)
