/*
 * CÓMO USARLO:
 * 1. Abre las herramientas de desarrollador (F12) en la página de iSAMS,
 *    pestaña "Console".
 * 2. Pega este script completo y presiona ENTER.
 * 3. Verás el mensaje "Listo. Haz clic en el elemento que quieres ubicar...".
 * 4. Haz clic normalmente en el elemento que te interesa (ej. un botón
 *    dentro de un shadow root). El clic funciona normal, no se bloquea.
 * 5. En la consola aparecerá la ruta exacta (selector por selector,
 *    incluyendo cada salto de shadow DOM) y también un pequeño snippet
 *    de Python listo para usar. Cópialo y pásamelo.
 * 6. Si quieres capturar otro elemento, vuelve a pegar el script.
 *
 * NOTA: usa event.composedPath() en vez de event.target, porque target se
 * "retargetea" al host cuando el elemento real está dentro de un shadow
 * root, incluso si es shadow root "open". composedPath() sí revela el
 * elemento real donde ocurrió el clic.
 */

(function () {
  function describir(el) {
    let sel = el.tagName.toLowerCase();
    if (el.id) {
      sel += "#" + el.id;
      return { selector: sel, unico: true };
    }
    if (el.className && typeof el.className === "string" && el.className.trim()) {
      const clases = el.className.trim().split(/\s+/).join(".");
      sel += "." + clases;
    }
    const padre = el.parentElement || (el.getRootNode() instanceof ShadowRoot ? el.getRootNode() : null);
    let indice = 0;
    if (padre) {
      const hermanosMismoTag = Array.from(padre.children || padre.querySelectorAll(":scope > *"))
        .filter((h) => h.tagName === el.tagName);
      indice = hermanosMismoTag.indexOf(el);
    }
    return { selector: sel, unico: false, indice: indice };
  }

  function manejarClic(evento) {
    document.removeEventListener("click", manejarClic, true);

    const ruta = evento.composedPath();
    // El primer elemento de composedPath() es el elemento real donde ocurrió el clic
    let nodo = ruta.find((n) => n.nodeType === 1); // primer Element real

    const saltos = [];

    while (nodo) {
      const raiz = nodo.getRootNode();
      const esShadow = raiz instanceof ShadowRoot;

      const cadenaLocal = [];
      let actual = nodo;
      while (actual && actual !== raiz) {
        cadenaLocal.unshift(describir(actual));
        actual = actual.parentElement;
      }
      saltos.unshift(cadenaLocal);

      if (esShadow) {
        nodo = raiz.host;
      } else {
        nodo = null;
      }
    }

    console.log("======= RUTA CAPTURADA (via composedPath) =======");
    console.log("Elemento real clicado:", ruta[0]);
    console.log("");
    console.log("--- Ruta legible ---");
    saltos.forEach((cadena, i) => {
      const texto = cadena
        .map((p) => (p.unico ? p.selector : `${p.selector} [índice ${p.indice}]`))
        .join(" > ");
      console.log(`Nivel ${i}${i > 0 ? " (dentro del shadow root anterior)" : ""}: ${texto}`);
    });

    console.log("");
    console.log("--- Snippet Python sugerido (ajusta si algún nivel no es único) ---");
    let codigo = "elemento = elemento_profundo(driver, " + JSON.stringify(construirSelectorSimple(saltos[0])) + ")\n";
    for (let i = 1; i < saltos.length; i++) {
      const selectorNivel = construirSelectorSimple(saltos[i]);
      codigo += `elemento = dentro_de_shadow(driver, elemento, ${JSON.stringify(selectorNivel)})\n`;
    }
    console.log(codigo);

    function construirSelectorSimple(cadena) {
      const ultimo = cadena[cadena.length - 1];
      return ultimo.selector;
    }

    alert("Ruta capturada (con composedPath). Revisa la consola (F12).");
  }

  document.addEventListener("click", manejarClic, true);
  console.log("Listo. Haz clic en el elemento que quieres ubicar...");
})();