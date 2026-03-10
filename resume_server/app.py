from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl import load_workbook
from copy import copy
import io
import os

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.xlsx')

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb['履歴書']

    def sc(addr, val):
        ws[addr] = str(val) if val else ''

    # 記入日
    from datetime import date
    today = date.today()
    sc('E3', f'　　{today.year}年{today.month}月{today.day}日現在')

    # 基本情報
    sc('C6', (data.get('lastKana','') + '　' + data.get('firstKana','')).strip())
    sc('C9', (data.get('lastName','') + '　' + data.get('firstName','')).strip())

    # 性別（先に書く）
    sc('N14', data.get('gender',''))

    # 生年月日
    by = data.get('birthYear','')
    bm = data.get('birthMonth','')
    bd = data.get('birthDay','')
    if by:
        from datetime import date as d
        try:
            birth = d(int(by), int(bm), int(bd))
            today2 = d.today()
            age = today2.year - birth.year - ((today2.month, today2.day) < (birth.month, birth.day))
            sc('B14', f"{by}年　{str(bm).zfill(2)}月　{str(bd).zfill(2)}日生　（満　{age}　歳）")
        except:
            sc('B14', f"{by}年　{bm}月　{bd}日生")

    # 住所
    sc('C16', data.get('addrKana',''))
    sc('H16', ' ' + data.get('phone',''))
    sc('C19', '〒' + data.get('zipCode','') + '　' + data.get('address',''))
    sc('H19', ' ' + data.get('email',''))

    # 学歴・職歴
    EDU_L = [35,38,41,44,47,50,53,56,59,62,65,68,71,74,77,80]
    EDU_R = [2,5,8,11,16,19]

    ent = []
    education = data.get('education', [])
    jobs = data.get('jobs', [])

    if education:
        ent.append({'yr':'','mo':'','txt':'学　歴'})
        for e in education:
            if e.get('enterYear'):
                ent.append({'yr':e['enterYear'],'mo':e.get('enterMonth',''),'txt':f"{e['name']}　入学"})
            if e.get('exitYear'):
                ent.append({'yr':e['exitYear'],'mo':e.get('exitMonth',''),'txt':f"{e['name']}　卒業"})

    if jobs:
        ent.append({'yr':'','mo':'','txt':'職　歴'})
        for j in jobs:
            if j.get('enterYear'):
                ent.append({'yr':j['enterYear'],'mo':j.get('enterMonth',''),'txt':f"{j['name']}　入社（{j.get('type','')}）"})
            if j.get('exitYear'):
                ent.append({'yr':j['exitYear'],'mo':j.get('exitMonth',''),'txt':f"{j['name']}　退社"})
            elif j.get('enterYear'):
                ent.append({'yr':'','mo':'','txt':'現在に至る'})

    ent.append({'yr':'','mo':'','txt':'以　上'})

    lft = ent[:16]
    rgt = ent[16:22]

    for i, r in enumerate(EDU_L):
        if i < len(lft):
            sc(f'B{r}', lft[i]['yr'])
            sc(f'C{r}', lft[i]['mo'])
            sc(f'D{r}', lft[i]['txt'])
        else:
            sc(f'B{r}', '')
            sc(f'C{r}', '')
            sc(f'D{r}', '')

    for i, r in enumerate(EDU_R):
        if i < len(rgt):
            sc(f'L{r}', rgt[i]['yr'])
            sc(f'M{r}', rgt[i]['mo'])
            sc(f'N{r}', rgt[i]['txt'])
        else:
            sc(f'L{r}', '')
            sc(f'M{r}', '')
            sc(f'N{r}', '')

    # 資格
    LIC_R = [25,28,31,34,37,40]
    licenses = data.get('licenses', [])
    for i, r in enumerate(LIC_R):
        if i < len(licenses):
            sc(f'L{r}', licenses[i].get('year',''))
            sc(f'M{r}', licenses[i].get('month',''))
            sc(f'N{r}', licenses[i].get('name',''))
        else:
            sc(f'L{r}', '')
            sc(f'M{r}', '')
            sc(f'N{r}', '')

    # 志望動機
    sc('L47', data.get('prText',''))

    # 本人希望欄
    hp = []
    if data.get('currentSalary'): hp.append(f"現在年収：{data['currentSalary']}万円")
    if data.get('desiredSalary'): hp.append(f"希望年収：{data['desiredSalary']}万円")
    if data.get('pcSkills'):      hp.append(f"PCスキル：{data['pcSkills']}")
    if data.get('hopeText'):      hp.append(data['hopeText'])
    sc('L71', '\n'.join(hp))

    # メモリ上に保存して返す
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    name = (data.get('lastName','') + data.get('firstName','')).strip() or '応募者'
    filename = f"履歴書_{name}.xlsx"

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
