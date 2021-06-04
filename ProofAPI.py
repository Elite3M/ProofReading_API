from flask import Flask
from flask_restful import Api, Resource, reqparse
import ProofAPI_functions

app = Flask(__name__)
api = Api(app)


class Proof(Resource):

    def get(self, dir):
        try:
            return {'data': ProofAPI_functions.main(dir)}
        except:
            return "Error processing document", 404


dir = "PEDTR2019F012"
api.add_resource(Proof, "/api/proof/<dir>")
if __name__ == '__main__':
    app.run(debug=True)
