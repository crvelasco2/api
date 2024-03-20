from flask import Flask, request, jsonify
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from flask_cors import CORS  # Importa la extensión Flask-CORS
import json

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "http://127.0.0.1:5000"}})
  # Habilita CORS en tu aplicación Flask

@app.route('/createItem', methods=['POST'])
def create_item():
    # Obtener los datos del cuerpo de la petición
    data = request.json
    usuario = data.get('Usuario')
    fecha = data.get('Fecha')
    unidad = data.get('Unidad')
    parada = data.get('Parada')
    novedad = data.get('Novedad')
    descripcionNovedad = data.get('DescripcionNovedad')

    # URL y credenciales de SharePoint
    site_url = "https://kreandotiec.sharepoint.com/sites/LaboratorioPRODSoftware/"
    username = "desarrollo03@kreandoti.com"
    password = "desaKreando032024"

    # Crear contexto de cliente con credenciales
    ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))

    # Obtener la lista por su título
    list_title = "Transporte Pesado"
    sp_list = ctx.web.lists.get_by_title(list_title)

    # Crear el payload para el ítem a ser creado
    item_properties = {
        'Usuario': usuario,
        'Fecha': fecha,
        'Unidad': unidad,
        'Parada': parada,
        'Noveda': novedad,
        'Descripci_x00f3_nNovedad': descripcionNovedad,  # Ajustado para coincidir con el nombre interno esperado
        }

    # Añadir el ítem a la lista
    list_item = sp_list.add_item(item_properties)
    ctx.execute_query()  # Ejecutar la solicitud para crear el ítem

    # Respuesta de éxito
    return jsonify({'message': 'Item created successfully', 'id': list_item.properties['Id']}), 201

if __name__ == '__main__':
    app.run(debug=True, port=5000)
