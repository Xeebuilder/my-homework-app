const fs = require('fs');
const https = require('https');

const key = Buffer.from('QVEuQWI4Uk42SnBsdlYxX0F3RzdJYWVfWWlLZFpDTFlidFBvRUV0WEVjX2pWczdFRFFPaXc=', 'base64').toString('utf-8');
const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=${key}`;

const entrada = `Origen del Universo y teoría del Big Bang
Estructura del Universo y la Vía Láctea
Componentes del Sistema Solar (Sol
planetas
satélites
asteroides
cometas
meteoritos y polvo cósmico)
Estructura del planeta Tierra y sus subsistemas (atmósfera
hidrosfera
geosfera y biosfera)
Ley de la gravedad en el Universo
Características y diferencias entre estrellas y planetas
Constelaciones del zodíaco
Medición de distancias astronómicas y años luz`;

const prompt = 
    "Actúa como un experto analizador de texto. Toma el texto proporcionado (que contiene temas y subtemas desordenados, con saltos de línea rotos y paréntesis) y conviértelo en una LISTA SEPARADA POR COMAS de temas individuales y completos.\n\n" +
    "INSTRUCCIONES ESTRICTAS:\n" +
    "1. Si un tema principal contiene subtemas entre paréntesis o en saltos de línea, EXTRAE cada subtema y conviértelo en un tema independiente que incluya el contexto del tema principal.\n" +
    "2. NO dejes paréntesis con pedazos de texto roto como 'meteoritos y polvo cósmico)'. Elimina los paréntesis integrando las palabras en un título lógico.\n" +
    "3. 'geosfera' y 'biosfera' DEBEN separarse en temas distintos aunque vengan juntos en el texto.\n" +
    "4. Cosas pequeñas y relacionadas como 'meteoritos' y 'polvo cósmico' pueden agruparse lógicamente en un solo tema.\n" +
    "5. Tu respuesta debe ser ÚNICAMENTE la lista separada por comas.\n\n" +
    "EJEMPLO DE ENTRADA:\n" +
    "Componentes del Sistema Solar (el Sol\n" +
    "planetas\n" +
    "asteroides\n" +
    "meteoritos y polvo cósmico)\n" +
    "Estructura del planeta Tierra (atmósfera\n" +
    "geosfera y biosfera)\n\n" +
    "EJEMPLO DE SALIDA ESPERADA:\n" +
    "El Sol como componente del Sistema Solar, Planetas del Sistema Solar, Asteroides del Sistema Solar, Meteoritos y polvo cósmico del Sistema Solar, Atmósfera del planeta Tierra, Geosfera del planeta Tierra, Biosfera del planeta Tierra\n\n" +
    "TEXTO A ANALIZAR:\n" + entrada;

const data = JSON.stringify({ contents: [{ role: 'user', parts: [{text: prompt}] }] });

const options = {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Content-Length': data.length
  }
};

const req = https.request(url, options, (res) => {
  let body = '';
  res.on('data', (d) => body += d);
  res.on('end', () => console.log(JSON.parse(body).candidates[0].content.parts[0].text));
});
req.write(data);
req.end();
