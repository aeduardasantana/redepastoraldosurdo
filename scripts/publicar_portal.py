"""Publicação exclusiva da Home, sem expor dados de gestão ou credenciais."""
import ftplib, io, os, re, secrets, ssl, sys, urllib.request
from html.parser import HTMLParser
from pathlib import Path

SITE = "https://redepastoraldosurdo.compassrosesystems.com.br/"
BASE = "public_html/redepastoraldosurdo"
PAGE = Path("publico/index.html")
CONTENT = PAGE.read_bytes()
class Audit(HTMLParser):
    def __init__(self):
        super().__init__()
        self.ids=set()
        self.anchors=[]
        self.h1=0
        self.images=[]
        self.scripts=[]
        self.lang=None
    def handle_starttag(self, tag, attrs):
        a=dict(attrs)
        if a.get("id"):
            assert a["id"] not in self.ids, "ID repetido"
            self.ids.add(a["id"])
        if tag=="html": self.lang=a.get("lang")
        if tag=="h1": self.h1+=1
        if tag=="a" and a.get("href","").startswith("#"): self.anchors.append(a["href"][1:])
        if tag=="img": self.images.append(a)
        if tag=="script": self.scripts.append(a.get("src"))
def audit():
    a=Audit()
    a.feed(CONTENT.decode("utf-8"))
    assert a.lang=="pt-BR" and a.h1==1
    assert all(x in a.ids for x in a.anchors)
    assert a.images and all(x.get("alt") for x in a.images)
    assert a.scripts==["https://vlibras.gov.br/app/vlibras-plugin.js"]
    assert b"data-portal-release=" in CONTENT and b"GEB Tecnologia" in CONTENT
    assert b"script.google.com/macros/s/" in CONTENT and b"pastoraldosurdocentro.blogspot.com" in CONTENT
    assert b"docs.google.com/spreadsheets" not in CONTENT and b"google.script.run" not in CONTENT
    assert b"<form" not in CONTENT and b"CPF_TOKEN" not in CONTENT
    print("Home: auditoria estrutural e isolamento da area privada aprovados.",flush=True)
class ImplicitTLS(ftplib.FTP_TLS):
    @property
    def sock(self): return self._sock
    @sock.setter
    def sock(self, value):
        if value is not None and not isinstance(value, ssl.SSLSocket):
            value=self.context.wrap_socket(value,server_hostname=self.host)
        self._sock=value
def get_https(path=""):
    url=SITE+path
    req=urllib.request.Request(url,headers={"Cache-Control":"no-cache","User-Agent":"PastoralPortalDeploy/1.0"})
    with urllib.request.urlopen(req,timeout=25,context=ssl.create_default_context()) as r:
        if r.status!=200 or r.geturl()!=url: raise RuntimeError("HTTPS divergente")
        return r.read(2_000_000)
def publish():
    host,user,password=(os.environ.get(x,"") for x in ("FTP_HOST","FTP_USER","FTP_PASSWORD"))
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9.-]*",host) or not user or not password:
        raise RuntimeError("Credenciais ausentes ou host invalido")
    # Nenhuma credencial é transmitida antes da validação TLS do servidor.
    ftp=None
    for mode,port,cls in [("explicito",21,ftplib.FTP_TLS),("implicito",990,ImplicitTLS)]:
        c=None
        try:
            c=cls(context=ssl.create_default_context(),timeout=15)
            c.connect(host,port)
            if mode=="explicito": c.auth()
            if not isinstance(c.sock,ssl.SSLSocket): raise RuntimeError("TLS ausente")
            print("FTPS "+mode+": certificado e identidade validados.",flush=True)
        except Exception as e:
            print("FTPS "+mode+": TLS nao confirmado ("+type(e).__name__+"). Nenhuma credencial enviada.",flush=True)
            if c: c.close()
            continue
        try:
            c.login(user,password)
            c.prot_p()
            c.cwd(BASE)
            ftp=c
            print("Conexao autenticada e pasta de destino acessivel.",flush=True)
            break
        except Exception as e:
            print("FTPS autenticado ou pasta nao confirmados ("+type(e).__name__+").",flush=True)
            c.close()
            break
    if ftp is None: raise RuntimeError("Transporte seguro indisponivel; nenhuma publicacao realizada")
    tmp="portal-release-"+secrets.token_hex(12)+".html"
    backup=None
    published=False
    staged=False
    try:
        names={x.rsplit("/",1)[-1] for x in ftp.nlst()}
        if "index.html" in names:
            buff=io.BytesIO()
            ftp.retrbinary("RETR index.html",buff.write)
            backup=buff.getvalue()
        # Primeiro comprovar correspondência do caminho FTP com o domínio.
        staged=True
        ftp.storbinary("STOR "+tmp,io.BytesIO(CONTENT))
        if get_https(tmp)!=CONTENT: raise RuntimeError("Arquivo temporario nao corresponde ao dominio")
        ftp.rename(tmp,"index.html")
        staged=False
        published=True
        if get_https("?release="+secrets.token_hex(8))!=CONTENT: raise RuntimeError("Home HTTP diferente do arquivo publicado")
        print("PUBLICACAO CONFIRMADA: index.html servido por HTTPS, conteudo identico ao GitHub.",flush=True)
    except Exception:
        if published:
            if backup is None:
                ftp.delete("index.html")
            else:
                restore="portal-restore-"+secrets.token_hex(12)+".html"
                ftp.storbinary("STOR "+restore,io.BytesIO(backup))
                ftp.rename(restore,"index.html")
            print("Reversao do index.html executada.",flush=True)
        raise
    finally:
        if staged:
            try: ftp.delete(tmp)
            except Exception: print("Limpeza temporaria nao confirmada.",flush=True)
        try: ftp.quit()
        except Exception: ftp.close()
if __name__=="__main__":
    try:
        audit()
        if "--publish" in sys.argv: publish()
    except Exception as e:
        # Não registrar respostas do servidor, credenciais ou conteúdo de arquivos remotos.
        print("::error::Etapa interrompida: "+(str(e) if isinstance(e,RuntimeError) else type(e).__name__),flush=True)
        sys.exit(1)
