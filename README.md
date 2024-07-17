<html lang="th">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>แบบสอบถาม SARC-F</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.16.9/xlsx.full.min.js"></script>
    <style>
        body {
            font-family: 'Arial', sans-serif;
            background-color: #ffe6f2;
            color: #333;
            margin: 0;
            padding: 0;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            box-sizing: border-box;
        }
        .container {
            background-color: #fff;
            border-radius: 10px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
            padding: 20px;
            max-width: 600px;
            width: 90%;
            margin: 20px;
            text-align: center;
        }
        h1 {
            color: #d63384;
            text-align: center;
        }
        .question {
            margin-bottom: 20px;
            text-align: left;
        }
        label {
            font-weight: bold;
            display: block;
            margin-bottom: 5px;
        }
        select {
            width: 100%;
            padding: 10px;
            margin-bottom: 10px;
            border: 2px solid #d63384;
            border-radius: 5px;
            box-sizing: border-box;
            font-size: 16px;
        }
        button {
            background-color: #d63384;
            color: #fff;
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            font-size: 16px;
            cursor: pointer;
            display: block;
            width: 100%;
            margin-top: 20px;
        }
        button:hover {
            background-color: #b3246d;
        }
        .result {
            margin-top: 20px;
            font-weight: bold;
            text-align: center;
            padding: 10px;
            border: 2px solid #d63384;
            border-radius: 5px;
            background-color: #ffe6f2;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>แบบสอบถามเพื่อคัดกรองภาวะมวลกล้ามเนื้อน้อยชนิด SARC-F</h1>

        <div class="question">
            <label for="q1">1. ท่านคิดว่าการยกและถือของที่มีน้ำหนัก 20 กิโลกรัมยากหรือไม่?</label>
            <select id="q1">
                <option value="0">ไม่ยาก</option>
                <option value="1">ยากเล็กน้อย</option>
                <option value="2">ยาก</option>
            </select>
        </div>

        <div class="question">
            <label for="q2">2. ท่านรู้สึกว่าการเดินภายในห้องยากหรือไม่?</label>
            <select id="q2">
                <option value="0">ไม่ยาก</option>
                <option value="1">ยากเล็กน้อย</option>
                <option value="2">ยาก</option>
            </select>
        </div>

        <div class="question">
            <label for="q3">3. ท่านเคลื่อนย้ายตัวเองจากเก้าอี้ไปเตียงนอนยากหรือไม่?</label>
            <select id="q3">
                <option value="0">ไม่ยาก</option>
                <option value="1">ยากเล็กน้อย</option>
                <option value="2">ยาก</option>
            </select>
        </div>

        <div class="question">
            <label for="q4">4. ท่านรู้สึกว่าการเดินขึ้นบันไดจำนวน 10 ขั้นยากหรือไม่?</label>
            <select id="q4">
                <option value="0">ไม่ยาก</option>
                <option value="1">ยากเล็กน้อย</option>
                <option value="2">ยาก</option>
            </select>
        </div>

        <div class="question">
            <label for="q5">5. ท่านเคยมีประวัติหกล้มมากี่ครั้ง?</label>
            <select id="q5">
                <option value="0">ไม่เคย</option>
                <option value="1">1-3 ครั้ง</option>
                <option value="2">มากกว่า 4 ครั้ง</option>
            </select>
        </div>

        <button onclick="calculateScore()">คำนวณคะแนน</button>

        <div class="result" id="result"></div>
        <button onclick="exportToExcel()">ดาวน์โหลดเป็น Excel</button>
    </div>

    <script>
        let totalScore = 0;
        let riskLevel = '';

        function calculateScore() {
            const q1 = parseInt(document.getElementById('q1').value);
            const q2 = parseInt(document.getElementById('q2').value);
            const q3 = parseInt(document.getElementById('q3').value);
            const q4 = parseInt(document.getElementById('q4').value);
            const q5 = parseInt(document.getElementById('q5').value);

            totalScore = q1 + q2 + q3 + q4 + q5;
            riskLevel = totalScore >= 4 ? "สูง" : "ต่ำ";

            document.getElementById('result').innerText = `คะแนนรวม SARC-F ของคุณคือ ${totalScore} ซึ่งบ่งชี้ว่าคุณมีความเสี่ยง ${riskLevel} ของการเกิดภาวะมวลกล้ามเนื้อน้อย.`;
        }

        function exportToExcel() {
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.json_to_sheet([
                { คำถาม: '1. ท่านคิดว่าการยกและถือของที่มีน้ำหนัก 20 กิโลกรัมยากหรือไม่?', คะแนน: document.getElementById('q1').value },
                { คำถาม: '2. ท่านรู้สึกว่าการเดินภายในห้องยากหรือไม่?', คะแนน: document.getElementById('q2').value },
                { คำถาม: '3. ท่านเคลื่อนย้ายตัวเองจากเก้าอี้ไปเตียงนอนยากหรือไม่?', คะแนน: document.getElementById('q3').value },
                { คำถาม: '4. ท่านรู้สึกว่าการเดินขึ้นบันไดจำนวน 10 ขั้นยากหรือไม่?', คะแนน: document.getElementById('q4').value },
                { คำถาม: '5. ท่านเคยมีประวัติหกล้มมากี่ครั้ง?', คะแนน: document.getElementById('q5').value },
                { คำถาม: 'คะแนนรวม', คะแนน: totalScore },
                { คำถาม: 'ระดับความเสี่ยง', คะแนน: riskLevel }
            ]);

            XLSX.utils.book_append_sheet(wb, ws, 'ผลการประเมิน');
            XLSX.writeFile(wb, 'sarc-f-assessment.xlsx');
        }
    </script>
</body>
</html>
