// Configuración compartida por index.html, registro.html, tarjeta.html y admin.html
const API_URL = 'https://script.google.com/macros/s/AKfycbzMznc0igGOlevzOmUx1h9J7u-kWu6N0FLbJg425F-cz64C09Dkx0WBbJIdjFtJEB2w/exec';

// Hashea el PIN en el navegador: nunca sale texto plano hacia la red.
async function sha256Hex(texto) {
  const buffer = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(texto));
  return Array.from(new Uint8Array(buffer)).map(b => b.toString(16).padStart(2, '0')).join('');
}
