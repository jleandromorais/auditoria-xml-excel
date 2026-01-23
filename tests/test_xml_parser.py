import tempfile
from pathlib import Path

from auditoria.xml_parser import parse_xml_file


def test_parse_nfe_minimo():
    xml = """<?xml version="1.0"?>
    <nfeProc xmlns="http://www.portalfiscal.inf.br/nfe">
      <NFe>
        <infNFe>
          <ide><nNF>123</nNF></ide>
          <total><ICMSTot>
            <vNF>100.00</vNF>
            <vICMS>10.00</vICMS>
            <vPIS>1.00</vPIS>
            <vCOFINS>2.00</vCOFINS>
          </ICMSTot></total>
        </infNFe>
      </NFe>
    </nfeProc>
    """
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "nfe.xml"
        p.write_text(xml, encoding="utf-8")
        info = parse_xml_file(str(p))
        assert info is not None
        assert info["Tipo"] == "NF-e"
        assert info["Nota"] == "123"
        assert info["Bruto"] == 100.0
        assert info["Liq_Calc"] == 87.0
