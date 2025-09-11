#!/usr/bin/env python3
import sys

def fix_indentation(filename):
    """Convertit tous les tabs en 4 espaces."""
    with open(filename, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Remplacer les tabs par 4 espaces
    content = content.replace('\t', '    ')
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"✅ Fichier {filename} corrigé")

if __name__ == '__main__':
    fix_indentation('amdec/utils/excel_handler.py')
