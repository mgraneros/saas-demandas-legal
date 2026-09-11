// frontend_demandas/auth-guard.js

// 1. CERRADO GLOBAL Y LIMPIEZA DE SESIÓN
function forzarCierreSesion(motivo = "Sesión finalizada.") {
  localStorage.removeItem("access_token");
  localStorage.removeItem("token");
  sessionStorage.clear();
  alert(motivo);
  window.location.href = "login.html";
}

// 2. INTERCEPTOR GLOBAL FETCH (HTTP 401)
const { fetch: originalFetch } = window;

window.fetch = async (...args) => {
  const response = await originalFetch(...args);

  // Si cualquier endpoint autenticado responde 401, el token venció
  if (response.status === 401) {
    const requestUrl = typeof args[0] === "string" ? args[0] : args[0]?.url || "";
    // Ignora peticiones que vayan a /token (el login maneja su propio error de credenciales)
    if (!requestUrl.includes("/token")) {
      forzarCierreSesion("Tu sesión ha expirado por seguridad. Ingresá nuevamente.");
    }
  }

  return response;
};

// 3. DETECTOR DE INACTIVIDAD (30 MINUTOS)
const TIEMPO_MAXIMO_INACTIVIDAD = 30 * 60 * 1000; // 30 minutos
let timeoutInactividad;

function reiniciarContadorInactividad() {
  clearTimeout(timeoutInactividad);

  // Solo corre si el usuario tiene una sesión iniciada
  const token = localStorage.getItem("access_token") || localStorage.getItem("token");
  if (!token) return;

  timeoutInactividad = setTimeout(() => {
    forzarCierreSesion("Tu sesión se cerró por 30 minutos de inactividad.");
  }, TIEMPO_MAXIMO_INACTIVIDAD);
}

// Eventos de interacción del usuario
["mousemove", "mousedown", "keydown", "touchstart", "scroll"].forEach((evento) => {
  window.addEventListener(evento, reiniciarContadorInactividad, { passive: true });
});

// Arranca el contador al cargar la vista
reiniciarContadorInactividad();