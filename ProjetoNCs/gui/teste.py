import requests

# URL do arquivo raw
version_file_url = "https://gitlab.egis-group.io/aa.silva/Hand_Helper_EGIS/-/raw/main/version.txt"
headers = {
    "Authorization": "Bearer 3cvR3Yxeni8S3_6nsHxS"  # Use seu token aqui
}

response = requests.get(version_file_url, headers=headers)

if response.status_code == 200:
    print("Conteúdo do arquivo:", response.text)
else:
    print(f"Erro ao acessar o arquivo: {response.status_code}")
