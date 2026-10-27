web: python -c "from database import apply_input_validation_migrations; apply_input_validation_migrations()" && python -m gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120
