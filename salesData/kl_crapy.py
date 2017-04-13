import requests
import json

r = requests.get('http://www.kaola.com/getFrontCategory.shtml')

def store_json(data):
	with open('category.json','w') as f:
		f.write(json.dumps(data))

