"""Deploy FTP convencional autorizado, restrito ao index.html público.
As credenciais percorrem FTP sem criptografia, como nos projetos anteriores.
Não publicar repositório, configurações, arquivos privados ou planilhas.
"""
import ftplib, io, os, re, secrets, sys
from publicar_portal import audit, CONTENT, BASE, get_https

def publish():
    host,user,password=(os.environ.get(x,"") for x in ("FTP_HOST","FTP_USER","FTP_PASSWORD"))
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9.-]*",host) or not user or not password:
        raise RuntimeError("Secrets HOST, USER, PASS ausentes ou host invalido.")
    ftp=ftplib.FTP(timeout=25)
    tmp="portal-release-"+secrets.token_hex(12)+".html"
    backup=None
    staged=False
    published=False
    try:
        ftp.connect(host,21)
        ftp.login(user,password)
        ftp.cwd(BASE)
        print("FTP convencional autenticado; diretorio exclusivo acessivel.",flush=True)
        names={x.rsplit("/",1)[-1] for x in ftp.nlst()}
        if "index.html" in names:
            buff=io.BytesIO()
            ftp.retrbinary("RETR index.html",buff.write)
            backup=buff.getvalue()
        staged=True
        ftp.storbinary("STOR "+tmp,io.BytesIO(CONTENT))
        if get_https(tmp)!=CONTENT:
            raise RuntimeError("Correspondencia entre FTP e dominio nao confirmada.")
        print("Arquivo temporario confirmado por HTTPS.",flush=True)
        ftp.rename(tmp,"index.html")
        staged=False
        published=True
        if get_https("?release="+secrets.token_hex(8))!=CONTENT:
            raise RuntimeError("Home servida difere da versao do GitHub.")
        print("PUBLICACAO CONFIRMADA: index.html identico ao GitHub, servido por HTTPS.",flush=True)
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
        publish()
    except Exception as e:
        print("::error::Etapa interrompida: "+(str(e) if isinstance(e,RuntimeError) else type(e).__name__),flush=True)
        sys.exit(1)
