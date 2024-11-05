import requests

# Substitua pelo seu token
github_token = "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"

# Teste a autenticação
headers = {
    "Authorization": f"token {github_token}"
}

response = requests.get("https://api.github.com/user", headers=headers)

if response.status_code == 200:
    print("Autenticação bem-sucedida!")
else:
    print(f"Falha na autenticação: {response.status_code} - {response.text}")
