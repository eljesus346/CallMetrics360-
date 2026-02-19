# test_imports.py
try:
    print("Intentando importar app.routers.reportes...")
    from app.routers.reportes import router
    print("✓ Importación exitosa!")
    print(f"Router: {router}")
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()