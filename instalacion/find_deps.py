import ast
import os
import sys
import sysconfig

def get_stdlib_modules():
    if sys.version_info >= (3, 10):
        return sys.stdlib_module_names
    # Fallback for older python
    return set()

def main(project_dir):
    stdlib = get_stdlib_modules()
    local_modules = set()
    for item in os.listdir(project_dir):
        if os.path.isdir(os.path.join(project_dir, item)):
            local_modules.add(item)
        elif item.endswith(".py"):
            local_modules.add(item[:-3])

    imports = set()
    
    for root, _, files in os.walk(project_dir):
        for file in files:
            if file.endswith('.py'):
                path = os.path.join(root, file)
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        tree = ast.parse(f.read(), filename=path)
                    
                    for node in ast.walk(tree):
                        if isinstance(node, ast.Import):
                            for name in node.names:
                                base_module = name.name.split('.')[0]
                                imports.add(base_module)
                        elif isinstance(node, ast.ImportFrom):
                            if node.module:
                                base_module = node.module.split('.')[0]
                                imports.add(base_module)
                except Exception:
                    pass

    third_party = []
    for imp in imports:
        if imp not in stdlib and imp not in local_modules and imp not in ['pip', 'setuptools', 'pkg_resources', 'pipreqs', 'setup']:
            third_party.append(imp)
            
    print("Posibles dependencias externas:")
    for imp in sorted(third_party):
        print("-", imp)

if __name__ == "__main__":
    main(r"c:\Users\eliseo_lopezp\proyecto1")
