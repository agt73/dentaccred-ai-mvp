import io
from app import app, analyze, chunks_from_text, SAMPLES, REVIEWS

def test_health():
    r=app.test_client().get('/health')
    assert r.status_code==200 and r.json['engine_mode']=='rules_assisted'

def test_seeded_statuses():
    expected={'supported':'Supported','partial':'Partially supported','insufficient':'Insufficient evidence','injection':'Expert review required'}
    for key,status in expected.items():
        cid,name,text=SAMPLES[key]
        review=analyze(cid,name,chunks_from_text(text))
        assert review['status']==status
        assert review['findings'] and review['notice']

def test_rejects_unsupported_upload():
    r=app.test_client().post('/reviews/new',data={'criterion':'2-8','period':'2025–2026','consent':'yes','document':(io.BytesIO(b'hello'),'evidence.txt')},content_type='multipart/form-data')
    assert r.status_code==200 and b'Only text-based PDF and DOCX' in r.data

def test_override_requires_comment():
    cid,name,text=SAMPLES['partial']; review=analyze(cid,name,chunks_from_text(text)); REVIEWS[review['id']]=review
    r=app.test_client().post(f"/reviews/{review['id']}/disposition",data={'action':'modify','status':'Potential concern','comment':''})
    assert r.status_code==302 and 'error=' in r.headers['Location']
