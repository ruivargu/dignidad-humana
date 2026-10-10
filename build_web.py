import docx,re,html,json,subprocess
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
import os; os.chdir(os.path.dirname(os.path.abspath(__file__))); SRC='SdA_Dignidad_Humana_3_4_ESO_v3.docx'
d=docx.Document(SRC)
els=list(d.element.body)
E=html.escape
def runs_html(p):
    out=[]
    for r in p.runs:
        t=E(r.text)
        if not t: continue
        if r.bold: t=f'<strong>{t}</strong>'
        if r.italic: t=f'<em>{t}</em>'
        out.append(t)
    return ''.join(out)
def clean_b(s): return re.sub(r'^•\s*','',s.replace('\t',' ')).strip()
def cell_html(tc):
    ps=[Paragraph(p,d) for p in tc.findall(qn('w:p'))]
    out=[];lis=[]
    def fl():
        nonlocal lis
        if lis: out.append('<ul>'+''.join(f'<li>{x}</li>' for x in lis)+'</ul>'); lis=[]
    for p in ps:
        if not p.text.strip(): continue
        h=runs_html(p)
        if p.text.startswith('•'): lis.append(clean_b(h))
        else: fl(); out.append(f'<p>{h}</p>')
    fl(); return ''.join(out)
def shaded(tc):
    shd=tc.find(qn('w:tcPr')+'/'+qn('w:shd'))
    return shd is not None and (shd.get(qn('w:fill')) or '').upper() not in ('AUTO','FFFFFF','')
def rows_text(tbl):
    out=[]
    for tr in tbl.findall(qn('w:tr')):
        out.append([cell_html(tc) for tc in tr.findall(qn('w:tc'))])
    return out
def plain(h): return re.sub(r'<[^>]+>','',h)
def rey_cards(tbl):
    rs=rows_text(tbl)[1:]; cards=[]
    for r in rs:
        m=re.match(r'(\d+)\.\s*(.*)',plain(r[0]))
        n,t=(m.group(1),m.group(2)) if m else ('',plain(r[0]))
        cards.append(f'<li class="pasaje"><span class="pn">{n}</span><h4><a href="#texto-rey-{n}">{E(t)}</a></h4><p class="idea">{plain(r[1])}</p><dl><div><dt>3º ESO</dt><dd>{plain(r[2])}</dd></div><div><dt>4º ESO</dt><dd>{plain(r[3])}</dd></div></dl></li>')
    return '<ol class="pasajes">'+''.join(cards)+'</ol>'
SHORT=('Fase','Duración','Dónde se hace','Pasaje de Rey','Pasajes de Rey')
def ficha_html(tbl):
    rs=rows_text(tbl)[1:]
    short=[];long_=[]
    for r in rs:
        lab=plain(r[0]).strip(); val=r[1]
        if not plain(val).strip(): continue
        (short if lab in SHORT else long_).append((lab,val))
    s=''.join(f'<div class="fact"><dt>{E(l)}</dt><dd>{v}</dd></div>' for l,v in short)
    l=''.join(f'<div class="long {slug(lab)}"><dt>{E(lab)}</dt><dd>{v}</dd></div>' for lab,v in long_)
    return f'<dl class="ficha2"><div class="facts">{s}</div>{l}</dl>'
def table_html(tbl):
    first=''.join(t.text or '' for t in tbl.findall(qn('w:tr'))[0].iter(qn('w:t')))
    if first.strip().startswith('Pasaje (bloque)'): return rey_cards(tbl)
    rows=tbl.findall(qn('w:tr'))
    grid=[]  # list of rows of dict(col,span,vm,tc)
    for tr in rows:
        c=0;row=[]
        for tc in tr.findall(qn('w:tc')):
            pr=tc.find(qn('w:tcPr'))
            gs=pr.find(qn('w:gridSpan')) if pr is not None else None
            span=int(gs.get(qn('w:val'))) if gs is not None else 1
            vm=pr.find(qn('w:vMerge')) if pr is not None else None
            vmv=None if vm is None else (vm.get(qn('w:val')) or 'continue')
            row.append(dict(col=c,span=span,vm=vmv,tc=tc)); c+=span
        grid.append(row)
    first=''.join(t.text or '' for t in rows[0].iter(qn('w:t')))
    cls='tbl'
    if first.strip().startswith('Ficha'): cls+=' ficha'
    out=[f'<div class="tw"><table class="{cls}">']
    for ri,row in enumerate(grid):
        cells=[]
        allth=all(shaded(c['tc']) for c in row if c['vm']!='continue')
        for c in row:
            if c['vm']=='continue': continue
            rs=1
            if c['vm']=='restart':
                for rj in range(ri+1,len(grid)):
                    m=[x for x in grid[rj] if x['col']==c['col']]
                    if m and m[0]['vm']=='continue': rs+=1
                    else: break
            tag='th' if shaded(c['tc']) else 'td'
            at=(f' colspan="{c["span"]}"' if c['span']>1 else '')+(f' rowspan="{rs}"' if rs>1 else '')
            cells.append(f'<{tag}{at}>{cell_html(c["tc"])}</{tag}>')
        out.append('<tr>'+''.join(cells)+'</tr>')
    out.append('</table></div>')
    return ''.join(out)
def mcls(m):
    m=m.lower()
    if m.startswith('tutor'): return 'tut'
    if m.startswith('educación'): return 'evce'
    if m.startswith('geograf'): return 'gh'
    if m.startswith('lengua'): return 'len'
    return 'tec'
def slug(s): 
    import unicodedata
    s=unicodedata.normalize('NFD',s.lower()); s=''.join(c for c in s if unicodedata.category(c)!='Mn')
    return re.sub(r'[^a-z0-9]+','-',s).strip('-')[:40]

# ---- parse into blocks
blocks=[]  # (kind,text/html,extra)
for e in els:
    if e.tag.endswith('}p'):
        p=Paragraph(e,d); t=p.text.strip()
        if not t: continue
        st=p.style.name
        if st.startswith('Heading'):
            blocks.append(('h'+st[-1],E(t),t)); continue
        h=runs_html(p)
        bold_all=all(r.bold for r in p.runs if r.text.strip()) and len(t)<95
        if t.startswith('•') : blocks.append(('li',clean_b(h),t))
        elif re.match(r'^\d+\.\t',p.text): blocks.append(('ol',re.sub(r'^\d+\.\s*','',h.replace('\t',' ')),t))
        elif bold_all: blocks.append(('h4',E(t),t))
        else: blocks.append(('p',h,t))
    elif e.tag.endswith('}tbl'):
        tt=''.join(t.text or '' for t in e.findall(qn('w:tr'))[0].iter(qn('w:t'))).strip()
        blocks.append(('tbl',table_html(e),'ficha' if tt.startswith('Ficha') else '',e))
# flatten lists
_TBL_N=[0]
def render(blks,big=False):
    out=[];lis=[];kind=None
    def fl():
        nonlocal lis,kind
        if lis: out.append(f'<{kind}>'+''.join(f'<li>{x}</li>' for x in lis)+f'</{kind}>'); lis=[];kind=None
    for b in blks:
        k=b[0]
        if k in('li','ol'):
            kk='ul' if k=='li' else 'ol'
            if kind and kind!=kk: fl()
            kind=kk; lis.append(b[1]); continue
        fl()
        if k=='h3': out.append(f'<h3>{b[1]}</h3>')
        elif k=='h4': out.append(f'<h4>{b[1]}</h4>')
        elif k=='p': out.append(f'<p>{b[1]}</p>')
        elif k=='tbl':
            tbl_html=b[1]
            if big and not tbl_html.startswith('<ol class="pasajes">'):
                rows=len(b[3].findall(qn('w:tr'))) if len(b)>3 else 0
                if rows>6:
                    _TBL_N[0]+=1; tid=f'tdlg-{_TBL_N[0]}'
                    tbl_html=(f'<button type="button" class="tabla-btn" data-dialog="{tid}" aria-haspopup="dialog">Ver tabla completa</button>'
                              f'<dialog class="tabla-dialog" id="{tid}"><div class="tabla-dialog-head"><strong>Tabla completa</strong>'
                              f'<button type="button" class="tabla-dialog-close" aria-label="Cerrar">×</button></div>'
                              f'<div class="tabla-dialog-body">{tbl_html}</div></dialog>')
            out.append(tbl_html)
    fl(); return '\n'.join(out)

title0=blocks[0][2]; h1=blocks[1][2]
rest=blocks[2:]
# split in H2 sections
secs=[];cur=None
for b in rest:
    if b[0]=='h2': cur={'title':b[2],'blocks':[]}; secs.append(cur)
    else: cur['blocks'].append(b)
NAV={'1':'Contexto','2':'Currículo','3':'Metodología','4':'Planificación','5':'Diversidad','6':'Evaluación','7':'Valoración','8':'Sesiones','Anexo I':'Anexo I','Anexo II':'Anexo II','Anexo III':'Anexo III'}

def fig(src,alt,w,h,cap):
    return f'<figure class="resource"><img src="img/{src}" alt="{E(alt)}" width="{w}" height="{h}" loading="lazy"><figcaption>{E(cap)}</figcaption></figure>'

def dcard(c,n,h,p):
    return f'<div class="dcard" style="--c:var(--{c})"><span class="dn">{n}</span><h5>{E(h)}</h5><p>{E(p)}</p></div>'
def step(c,n,h,p,date=None):
    d=f'<p class="sdate">{E(date)}</p>' if date else ''
    return f'<div class="step" style="--c:var(--{c})"><span class="sn">{n}</span>{d}<h5>{E(h)}</h5><p>{E(p)}</p></div>'
def diagram(inner,cap):
    return f'<figure class="resource diagram">{inner}<figcaption>{E(cap)}</figcaption></figure>'

IMG_SITUACION=diagram('<div class="dcards">'+
  dcard('w1','1','Aprendizaje basado en proyectos','Producto final con impacto en la comunidad educativa.')+
  dcard('w2','2','Producto real','Elaboración de un vídeo que recoge todo lo trabajado en las distintas materias.')+
  dcard('w3','3','Participación en concursos','Trabajo cooperativo en torno a valores democráticos: la dignidad humana.')+
  dcard('marker','4','Tutoría','Motor y seguimiento del proyecto a lo largo del curso.')+
  '</div>',"Por qué este tipo de proyecto, con producto real y proyección en un concurso, refuerza el aprendizaje.")

IMG_ESTRUCTURA=diagram('<div class="stepper">'+
  step('w1','1','Motivación','¿Qué sabemos? Arranque con el pasaje de Rey Martínez y la pregunta inicial.')+
  step('w2','2','Desarrollo','Diez sesiones de trabajo por materia, con preparación y producto parcial en cada una.')+
  step('w3','3','Producto final','Grabación y montaje del vídeo de dos minutos entre el alumnado.')+
  step('marker','4','Difusión','Presentación al X Premio Escolar Valores Democráticos.')+
  '</div>',"Las cuatro fases por las que pasa el proyecto, de la motivación inicial a la difusión del vídeo.")

IMG_FASES=diagram('<div class="stepper">'+
  step('w1','1','Entrega de equipos','Los equipos terminan el vídeo fuera del aula, con la Tutoría y el docente responsable como apoyo.',"Hasta finales de marzo de 2027")+
  step('w2','2','Selección interna','Participación del alumnado y comprobación de que las autorizaciones de imagen y voz de las familias están firmadas.',"Primera quincena de abril (sesión 10)")+
  step('w3','3','Envío de los vídeos','El docente responsable de cada equipo envía los vídeos elegidos, según las bases definitivas.',"Antes del 25 de abril de 2027")+
  step('marker','4','Calificación con rúbrica','Todos los vídeos se califican con la rúbrica, se presenten o no al concurso.',"Evaluación general")+
  '</div>',"Fechas clave del proyecto, desde la grabación hasta la evaluación final con rúbrica.")

IMG_USO=diagram('<div class="guia-uso"><div><h5 class="guia-h">Estructura de cada sesión</h5>'+
  '<ol class="guia-steps">'+
  '<li><strong>Ficha de la sesión</strong></li>'+
  '<li><strong>Preparación previa y materiales</strong></li>'+
  '<li><strong>Desarrollo por momentos</strong> (con minutos asignados)</li>'+
  '<li><strong>Voz del docente y del alumnado</strong></li>'+
  '<li><strong>Producto y evidencia de evaluación</strong></li>'+
  '<li><strong>Variantes</strong> (si falta o sobra tiempo)</li>'+
  '</ol><p class="guia-note">* Las sesiones 2 a 9 se apoyan en el texto de Fernando Rey Martínez (lectura por pasajes, sección 1).</p></div>'+
  '<div><h5 class="guia-h">Cómo leer esta guía</h5><dl class="guia-legend">'+
  '<div style="--c:var(--w1)"><dt>Sesiones comunes <span>(1, 4, 5, 6, 7 y 10)</span></dt><dd>Comunes a 3º y 4º ESO; cada ficha especifica los indicadores de cada curso.</dd></div>'+
  '<div style="--c:var(--w2)"><dt>Sesiones específicas <span>(2 y 3)</span></dt><dd>Versión propia para 3º y para 4º.</dd></div>'+
  '<div style="--c:var(--w3)"><dt>Sesiones de grabación <span>(8 y 9)</span></dt><dd>Se adaptan según la materia o el espacio.</dd></div>'+
  '<div style="--c:var(--marker)"><dt>Códigos de indicadores</dt><dd>Se desarrollan en el Anexo I (3º) y el Anexo II (4º).</dd></div>'+
  '</dl></div></div>',"Cómo está organizada cada sesión y qué tipo de ficha vas a encontrar en cada una.")

IMG_ANEXOS=diagram('<div class="anexo-flow"><div class="anexo-box main"><h5>Documento principal</h5><p>La situación de aprendizaje completa: contexto, metodología, planificación y guía del docente.</p></div><span class="anexo-arrow" aria-hidden="true">→</span><div class="anexo-list">'+
  '<div class="anexo-box"><h5>Anexo I · 3º ESO</h5><p>Fundamentación curricular e indicadores de logro.</p></div>'+
  '<div class="anexo-box"><h5>Anexo II · 4º ESO</h5><p>Fundamentación curricular e indicadores de logro.</p></div>'+
  '<div class="anexo-box"><h5>Anexo III</h5><p>Criterios, instrumentos y rúbrica de evaluación.</p></div>'+
  '<div class="anexo-box"><h5>Anexo IV</h5><p>Textos de referencia completos, de Rey Martínez y de Gil-Robles.</p></div>'+
  '</div></div>',"Los anexos amplían el documento principal con las tablas de indicadores y los materiales de apoyo.")
DIAG_SECUENCIA=diagram('<dl class="ficha2" style="--c:var(--w1)"><div class="facts">'+
  '<div class="fact"><dt>Duración</dt><dd>Uno o dos trimestres, según el horario del centro</dd></div>'+
  '<div class="fact"><dt>Plazo del concurso</dt><dd>Envío de los vídeos antes del 25 de abril de 2027</dd></div>'+
  '</div></dl>',"La temporalización se ajusta al horario real del grupo, dejando tiempo para terminar el producto final antes del plazo del concurso.")

DIAG_MATERIAS=diagram('<p style="margin:4px 0 12px"><span class="chip">3º ESO</span><span class="chip">4º ESO</span><span class="chip">Tutoría</span></p>'+
  '<p style="margin:0;max-width:52ch;color:var(--muted);font-size:.92rem">Las diez sesiones de cada curso se reparten entre las materias del itinerario, con la Tutoría como hilo conductor en ambos (detalle completo en «Recorrido de las 10 sesiones», al inicio de la guía).</p>',
  "Las mismas sesiones conectan con 3º ESO, 4º ESO y Tutoría.")

IMG_PLANIF_ROW='<div class="resource-row">'+DIAG_SECUENCIA+DIAG_MATERIAS+'</div>'
def navkey(t):
    m=re.match(r'(\d+)\.',t)
    if m: return m.group(1)
    m=re.match(r'(Anexo [IV]+)\.',t)
    return m.group(1) if m else t
sections_html=[];nav=[]
def sec_id(t):
    k=navkey(t); return 'anexo-'+k.split()[1].lower() if k.startswith('Anexo') else 'sec-'+k

def tabs_for_planning(blks):
    # group h3 containing 3º / 4º ESO + following table into tabs
    out=[];i=0;tabs=[]
    res=[]
    while i<len(blks):
        b=blks[i]
        if b[0]=='h3' and re.match(r'4\.[12] Itinerario de [34]º ESO',b[2]) and i+1<len(blks) and blks[i+1][0]=='tbl':
            tabs.append((b[2],blks[i+1][1])); i+=2; continue
        res.append(b); i+=1
    return res,tabs

def render_section(sec):
    t=sec['title']; k=navkey(t); sid=sec_id(t)
    nav.append((sid,NAV.get(k,t)))
    blks=sec['blocks']
    body=''
    if k=='4':
        res,tabs=tabs_for_planning(blks)
        # place tabs after first paragraph
        pre=[];post=[]
        placed=False
        for b in res:
            (pre if not placed else post).append(b)
            if not placed and b[0]=='p': placed=True
        tb='<div class="tabs" role="tablist" aria-label="Itinerario">'+''.join(f'<button role="tab" id="tab4-{i}" aria-selected="{"true" if i==0 else "false"}" aria-controls="pan4-{i}" data-tab="{i}">{E(x[0][4:])}</button>' for i,x in enumerate(tabs))+'</div>'
        pans=''.join(f'<div class="panel" role="tabpanel" id="pan4-{i}" aria-labelledby="tab4-{i}" {"hidden" if i else ""}>{x[1]}</div>' for i,x in enumerate(tabs))
        body=render(pre,True)+IMG_PLANIF_ROW+tb+pans+render(post,True)
    elif k=='8':
        intro=[];sess=[];cs=None
        for b in blks:
            if b[0]=='h3' and b[2].startswith('Sesión '):
                cs={'t':b[2],'b':[]}; sess.append(cs)
            elif cs is None: intro.append(b)
            else: cs['b'].append(b)
        body=render(intro)+IMG_USO
        body+='''<div class="filter" role="group" aria-label="Filtrar sesiones por curso"><span>Ver sesiones de</span><button data-f="all" aria-pressed="true">Todas</button><button data-f="3" aria-pressed="false">3º ESO</button><button data-f="4" aria-pressed="false">4º ESO</button><button class="ghost" id="openall">Abrir todas</button></div>'''
        cards=[]
        for i,s in enumerate(sess):
            m=re.match(r'Sesión (\d+)(?: \((3º ESO|4º ESO)\))?\. (.*?): (.*)$',s['t'])
            n,curso,mat,sub=m.groups(); sub=re.sub(r' \(3º y 4º\)$','',sub); sub=sub[0].upper()+sub[1:]
            cur_attr={'3º ESO':'3','4º ESO':'4'}.get(curso,'34')
            chip=curso or '3º y 4º ESO'
            fb=s['b']; ti=next((j for j,b in enumerate(fb) if b[0]=='tbl'),0); fb_tbl={ti:fb[ti][3]}
            ficha=ficha_html(fb_tbl[ti]); more=render(fb[ti+1:])
            csid=f's{n}'+('-3' if curso=='3º ESO' else '-4' if curso=='4º ESO' else '')
            pdf='fichas/sesiones/Sesion_'+n.zfill(2)+('_3ESO' if curso=='3º ESO' else '_4ESO' if curso=='4º ESO' else '')+'.pdf'
            dl=f'<a class="dl" href="{pdf}" target="_blank" rel="noopener">Descargar PDF</a>'
            cards.append(f'''<article class="sesion m-{mcls(mat)}" data-c="{cur_attr}" id="{csid}"><header><span class="num">{n}</span><div><p class="meta"><span class="chip">{chip}</span> {E(mat)}</p><h3>{E(sub)}</h3></div>{dl}</header>{ficha}<details><summary>Guía completa de la sesión</summary><div class="more">{more}</div></details></article>''')
        body+='<div class="sesiones">'+''.join(cards)+'</div>'
    elif k=='Anexo I' or k=='Anexo II' or k=='Anexo III':
        intro=[];ann=[];cs=None
        for b in blks:
            if b[0]=='h3': cs={'t':b[1],'b':[]}; ann.append(cs)
            elif cs is None: intro.append(b)
            else: cs['b'].append(b)
        body=render(intro,True)+''.join(f'<details class="anx"><summary>{a["t"]}</summary>{render(a["b"])}</details>' for a in ann)
        if k=='Anexo I': body=IMG_ANEXOS+body
    elif k=='1':
        body=IMG_SITUACION+render(blks,True)
    elif k=='3':
        body=IMG_ESTRUCTURA+render(blks,True)
    elif k=='7':
        body=IMG_FASES+render(blks,True)
    else:
        body=render(blks,True)
    return f'<section id="{sid}"><h2><span>{E(t)}</span></h2>{body}</section>'
secs_out=[render_section(s) for s in secs]
# the "Anexo III" lists h3 as plain h3; fine
# nav: shorten to those present
navhtml='<a href="#recorrido">Recorrido</a>'+''.join(f'<a href="#{i}">{E(l)}</a>' for i,l in nav)
main=''.join(secs_out)
NAMES={'tut':'Tutoría','evce':'Educación en Valores Cívicos y Éticos','gh':'Geografía e Historia','len':'Lengua Castellana y Literatura','tec':'Tecnología y Digitalización'}
T3=[('s1','tut'),('s2-3','evce'),('s3','gh'),('s4','evce'),('s5','len'),('s6','len'),('s7','tut'),('s8','tec'),('s9','tec'),('s10','tut')]
T4=[('s1','tut'),('s2-4','gh'),('s3','gh'),('s4','gh'),('s5','len'),('s6','len'),('s7','tut'),('s8','tec'),('s9','tec'),('s10','tut')]
def track(lbl,items,extra=''):
    nodes=''.join(f'<li><a href="#{sid}" class="node m-{c}" data-open="{sid}" title="Sesión {i+1}: {NAMES[c]}"><span class="sr">Sesión </span>{i+1}</a></li>' for i,(sid,c) in enumerate(items))
    return f'<div class="row"><h3>{lbl}</h3><ol class="track">{nodes}</ol>{extra}</div>'
legend=''.join(f'<li><i class="sw m-{k}"></i>{v}</li>' for k,v in NAMES.items())
TIMELINE=f'<section id="recorrido" class="reco"><h2><span>Recorrido de las 10 sesiones</span></h2><p>Cada sesión se enlaza con su ficha. En 4º ESO las sesiones 8 y 9 se hacen en Tecnología o en la hora de Tutoría, según el centro.</p>{track("3º ESO",T3)}{track("4º ESO",T4)}<ul class="legend">{legend}</ul></section>'

# ---- textos originales (Rey y Gil-Robles), texto completo + bloques con ancla
def _doc_lines(path):
    r=subprocess.run(['textutil','-convert','txt','-stdout',path],capture_output=True,text=True,check=True)
    return [l.strip() for l in r.stdout.split('\n') if l.strip() and not l.strip().startswith('PAGE')]
def ptag(paras): return ''.join(f'<p>{E(x)}</p>' for x in paras)
def bloque_art(prefix,n,titulo,paras):
    return f'<details class="bloque-texto" id="{prefix}-{n}"><summary><span class="bn">{n}</span> {E(titulo)}</summary><div class="texto-contenido">{ptag(paras)}</div></details>'

rey_lines=_doc_lines('doc/dignidad2_FRey.doc')
REY_BODY=rey_lines[2:22]
REY_BIO=rey_lines[22] if len(rey_lines)>22 else ''
REY_BLOCKS=[
 (1,'Kant y la persona como fin',REY_BODY[0:1]),
 (2,'Reconocimiento jurídico tras la Segunda Guerra Mundial y dignidad social',REY_BODY[1:2]),
 (3,'El artículo 10.1 de la Constitución',REY_BODY[2:3]),
 (4,'Las tres dificultades del concepto',REY_BODY[3:7]),
 (5,'Para qué sirve el concepto en la práctica',REY_BODY[7:8]),
 (6,'Las tres dimensiones de la dignidad',REY_BODY[8:12]),
 (7,'Dignidad social y mínimo existencial',REY_BODY[13:14]),
 (8,'Posibles desarrollos de la idea de dignidad para el alumnado',REY_BODY[14:20]),
]
rey_blocks_html=''.join(bloque_art('texto-rey',n,t,p) for n,t,p in REY_BLOCKS)
rey_full=ptag(REY_BODY)+f'<p class="credito">{E(REY_BIO)}</p>'

gd=docx.Document('doc/Principios y valores en crisis-corregido Alvaro.docx')
gr_lines=[p.text.strip() for p in gd.paragraphs if p.text.strip()]
GR_PRE=gr_lines[2]
GR_BLOCKS=[
 ('I','El largo camino histórico hacia el reconocimiento de la dignidad',gr_lines[3:10]),
 ('II','La Declaración Universal y el proyecto europeo',gr_lines[10:22]),
 ('III','Compartir y defender los valores: el compromiso constitucional',gr_lines[22:31]),
]
gr_blocks_html=''.join(bloque_art('texto-gr',n,t,p) for n,t,p in GR_BLOCKS)
gr_full=ptag([GR_PRE]+GR_BLOCKS[0][2]+GR_BLOCKS[1][2]+GR_BLOCKS[2][2])

TEXTOS=f'''<section id="anexo-iv"><h2><span>Anexo IV. Textos de referencia</span></h2>
<p>Este anexo reproduce, íntegros y sin cambios, los dos textos de referencia de la SdA. Los bloques temáticos con los que trabaja la guía son ocho pasajes en el texto de Rey (Anexo IV.A) y tres fragmentos en el de Gil-Robles (Anexo IV.B). Cada sesión parte de un pasaje o fragmento concreto; aquí puedes leerlos enteros o ir directamente al bloque citado en cada ficha.</p>
<div class="autor">
<h3>IV.A. Fernando Rey Martínez · «Dignidad humana»</h3>
<p class="credito">Catedrático de Derecho Constitucional de la Universidad de Valladolid. Texto guía: todas las sesiones parten de uno de sus ocho pasajes.</p>
<details class="texto-completo"><summary>Leer el texto completo</summary>{rey_full}</details>
<div class="bloques">{rey_blocks_html}</div>
</div>
<div class="autor">
<h3>IV.B. Álvaro Gil-Robles y Gil Delgado · «Principios y valores en crisis»</h3>
<p class="credito">Texto de apoyo secundario y puntual. Esta guía cita los fragmentos I y III; el fragmento II queda aquí disponible para quien quiera ampliarlo.</p>
<details class="texto-completo"><summary>Leer el texto completo</summary>{gr_full}</details>
<div class="bloques">{gr_blocks_html}</div>
</div>
</section>'''

navhtml+='<a href="#anexo-iv">Anexo IV</a>'
main+=TEXTOS
tpl=open('tpl.html').read()
out=tpl.replace('{{NAV}}',navhtml).replace('{{TITLE0}}',E(title0)).replace('{{H1}}',E(h1)).replace('{{MAIN}}',TIMELINE+main).replace('{{SIDE}}',navhtml)
open('index.html','w').write(out)
print(len(out)//1024,'KB',[s['title'][:30] for s in secs])
