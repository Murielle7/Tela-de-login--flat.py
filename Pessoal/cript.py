<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema de Criptografia e Descriptografia</title>
    <style>
        body {
            margin: 0;
            padding: 0;
            font-family: Arial, sans-serif;
            color: white;
            overflow: auto;
            background-color: black;
        }

        #matrix {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: url('https://www.example.com/path-to-high-resolution-matrix-image.jpg') no-repeat center center fixed;
            background-size: cover; /* Ajusta a imagem para cobrir toda a área da tela */
            z-index: -1;
        }

        #main {
            max-width: 800px;
            margin: 50px auto;
            padding: 20px;
            background: rgba(0, 0, 0, 0.7);
            border-radius: 8px;
            box-shadow: 0 0 10px rgba(16, 140, 5, 0.8);
            z-index: 10;
            position: relative;
            text-align: center;
            overflow: auto;
        }

        h1 {
            text-align: center;
            color: #ffffff;
        }

        textarea, input[type="text"] {
            width: 100%;
            padding: 10px;
            margin: 10px 0;
            border: 1px solid #333;
            border-radius: 4px;
            box-sizing: border-box;
        }

        button {
            background-color: #0c6d23;
            color: #fff;
            border: none;
            padding: 10px 20px;
            margin: 5px;
            border-radius: 4px;
            cursor: pointer;
        }

        button:hover {
            background-color: #074d17;
        }

        .hidden {
            display: none;
        }

        #result {
            margin-top: 20px;
        }
    </style>
</head>
<body>

    <canvas id="matrix"></canvas>

    <div id="main">
        <h1>Sistema de Criptografia e Descriptografia</h1>
        <div id="main-content">
            <button onclick="showCryptography()">Criptografia</button>
            <button onclick="showDecryption()">Descriptografia</button>
        </div>
        <div id="cryptography" class="hidden">
            <textarea id="text-to-encrypt" placeholder="Digite o texto para criptografar"></textarea>
            <input type="text" id="encryption-password" placeholder="Digite a senha">
            <button onclick="encryptText()">Criptografar</button>
            <button onclick="backToMain()">Voltar</button>
        </div>
        <div id="decryption" class="hidden">
            <textarea id="text-to-decrypt" placeholder="Digite o texto criptografado"></textarea>
            <input type="text" id="decryption-password" placeholder="Digite a senha">
            <button onclick="decryptText()">Descriptografar</button>
            <button onclick="backToMain()">Voltar</button>
        </div>
        <div id="result" class="hidden">
            <h2>Resultado</h2>
            <a id="download-btn" class="hidden" download>Baixar Arquivo</a>
        </div>
    </div>

    <script>
        const canvas = document.getElementById('matrix');
        const ctx = canvas.getContext('2d');
        canvas.width = window.innerWidth;
        canvas.height = window.innerHeight;

        const matrixChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ123456789@#$%^&*()*&^%";
        const fontSize = 16;
        const columns = canvas.width / fontSize; 
        const drops = Array(Math.floor(columns)).fill(1); 

        function drawMatrix() {
            ctx.fillStyle = "rgba(0, 0, 0, 0.05)";
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            ctx.fillStyle = "#0F0"; 
            ctx.font = fontSize + "px monospace";
            drops.forEach((y, index) => {
                const text = matrixChars.charAt(Math.floor(Math.random() * matrixChars.length));
                const x = index * fontSize;
                ctx.fillText(text, x, y * fontSize);
                if (y * fontSize > canvas.height && Math.random() > 0.975) {
                    drops[index] = 0;
                }
                drops[index]++;
            });
        }
        setInterval(drawMatrix, 50);

        function xorEncryptDecrypt(text, password) {
            const key = Array.from(password).map(c => c.charCodeAt(0)).reduce((a, b) => a + b, 0);
            let result = '';
            for (let i = 0; i < text.length; i++) {
                result += String.fromCharCode(text.charCodeAt(i) ^ key);
            }
            return result;
        }

        function encryptText() {
            const password = document.getElementById('encryption-password').value;
            const text = document.getElementById('text-to-encrypt').value;

            if (!text || !password) {
                alert("Por favor, preencha o texto e a senha.");
                return;
            }

            const encrypted = xorEncryptDecrypt(text, password);
            const blob = new Blob([encrypted], { type: 'text/plain' });

            const url = URL.createObjectURL(blob);
            const link = document.getElementById('download-btn');
            link.href = url;
            link.download = 'encrypted.txt';
            link.classList.remove('hidden');
            link.innerText = 'Baixar Arquivo';
            document.getElementById('result').classList.remove('hidden');
        }

        function decryptText() {
            const password = document.getElementById('decryption-password').value;
            const text = document.getElementById('text-to-decrypt').value;

            if (!text || !password) {
                alert("Por favor, preencha o texto criptografado e a senha.");
                return;
            }

            const decrypted = xorEncryptDecrypt(text, password);
            const blob = new Blob([decrypted], { type: 'text/plain' });

            const url = URL.createObjectURL(blob);
            const link = document.getElementById('download-btn');
            link.href = url;
            link.download = 'decrypted.txt';
            link.classList.remove('hidden');
            link.innerText = 'Baixar Arquivo';
            document.getElementById('result').classList.remove('hidden');
        }

        function showCryptography() {
            document.getElementById('main-content').classList.add('hidden');
            document.getElementById('cryptography').classList.remove('hidden');
            document.getElementById('decryption').classList.add('hidden');
            document.getElementById('result').classList.add('hidden');
        }

        function showDecryption() {
            document.getElementById('main-content').classList.add('hidden');
            document.getElementById('cryptography').classList.add('hidden');
            document.getElementById('decryption').classList.remove('hidden');
            document.getElementById('result').classList.add('hidden');
        }

        function backToMain() {
            document.getElementById('main-content').classList.remove('hidden');
            document.getElementById('cryptography').classList.add('hidden');
            document.getElementById('decryption').classList.add('hidden');
            document.getElementById('result').classList.add('hidden');
        }
    </script>
</body>
</html>
