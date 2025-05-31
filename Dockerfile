FROM regis.tjba.jus.br/k8s/imagens/python/python3.11:1.0.0

COPY extrator_dgt /extrator_dgt
COPY requirements.txt /tmp/requirements.txt
COPY extrator_dgt/settings_base.py /extrator_dgt/settings.py
RUN pip3 install -r /tmp/requirements.txt && chmod +x /extrator_dgt/servicos.sh && cd /

# Instantclient
RUN apt-get update -y && apt-get install -y --no-install-recommends fakeroot alien libaio1 && fakeroot alien --scripts --install /extrator_dgt/vendor/Oracle/RPM/*.rpm

WORKDIR /
CMD ["/extrator_dgt/servicos.sh"]
