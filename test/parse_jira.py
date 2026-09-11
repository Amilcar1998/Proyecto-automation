import json

with open(r'C:\Users\eliseo_lopezp\proyecto1\test\jira_meta.json', encoding='utf-8') as f:
    data = json.load(f)

for project in data.get('projects', []):
    if project.get('key') == 'ML':
        for issue_type in project.get('issuetypes', []):
            if issue_type.get('name') == 'Tarea':
                fields = issue_type.get('fields', {})
                for k, v in fields.items():
                    name = v.get('name', '')
                    if 'Aplication' in name or 'Application' in name or 'customfield' in k:
                        allowed = []
                        for av in v.get('allowedValues', []):
                            av_id = av.get('id', '')
                            av_val = av.get('value', '')
                            allowed.append(f"{av_val} (ID: {av_id})")
                        print(f"Field: {k} | Name: {name} | Options: {allowed}")
