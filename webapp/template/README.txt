Place your template DXF file here as:

    template.dxf

It will be mounted read-only into the Docker container at /app/template/template.dxf
and used automatically for every generation run.

To use a different filename or path, set the TEMPLATE_DXF environment variable
in docker-compose.yml:

    environment:
      - TEMPLATE_DXF=/app/template/your_custom_name.dxf
