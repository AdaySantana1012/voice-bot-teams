/**
 * voicebot-extensions.js
 *
 * 1) Registra un custom element "Input.Rating" para la librería adaptivecards.
 *    Pinta 5 estrellas clicables con estilo Bankinter (naranja #FF8D30).
 *
 * 2) Conecta markdown-it con la librería adaptivecards para que los TextBlocks
 *    que vengan con Markdown se rendericen correctamente.
 *
 * 3) Exporta una función parseMarkdown(text) para usar en burbujas normales del chat.
 *
 * Cargar DESPUÉS de adaptivecards.min.js y markdown-it.min.js
 */

// ============ 1) Custom Input.Rating ============

class InputRating extends AdaptiveCards.Input {

  // Propiedades del schema
  static idProperty = new AdaptiveCards.StringProperty(AdaptiveCards.Versions.v1_0, "id", true);
  static colorProperty = new AdaptiveCards.StringProperty(AdaptiveCards.Versions.v1_0, "color");

  _selectedValue = 0;
  _starElements = [];
  _rootElement = null;

  // Render: crea 5 estrellas SVG clicables
  internalRender() {
    this._rootElement = document.createElement("div");
    this._rootElement.style.display = "flex";
    this._rootElement.style.gap = "4px";
    this._rootElement.style.padding = "8px 0";
    this._rootElement.style.cursor = "pointer";
    this._rootElement.setAttribute("role", "radiogroup");
    this._rootElement.setAttribute("aria-label", "Valoración");

    this._starElements = [];

    for (let i = 1; i <= 5; i++) {
      const star = document.createElement("span");
      star.innerHTML = this._starSVG(false);
      star.style.transition = "transform 0.15s";
      star.setAttribute("role", "radio");
      star.setAttribute("aria-checked", "false");
      star.setAttribute("aria-label", `${i} estrella${i > 1 ? 's' : ''}`);
      star.setAttribute("tabindex", "0");

      star.addEventListener("click", () => this._setRating(i));
      star.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          this._setRating(i);
        }
      });
      star.addEventListener("mouseenter", () => this._highlightStars(i));
      star.addEventListener("mouseleave", () => this._highlightStars(this._selectedValue));

      this._starElements.push(star);
      this._rootElement.appendChild(star);
    }

    return this._rootElement;
  }

  _setRating(value) {
    this._selectedValue = value;
    this._highlightStars(value);
    this._starElements.forEach((s, idx) => {
      s.setAttribute("aria-checked", idx < value ? "true" : "false");
    });
    this.valueChanged();
  }

  _highlightStars(count) {
    this._starElements.forEach((star, idx) => {
      star.innerHTML = this._starSVG(idx < count);
      star.style.transform = idx < count ? "scale(1.15)" : "scale(1)";
    });
  }

  _starSVG(filled) {
    const color = filled ? "#FF8D30" : "#D0D0D0";
    return `<svg width="28" height="28" viewBox="0 0 24 24" fill="${filled ? color : 'none'}" stroke="${color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
      <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>
    </svg>`;
  }

  // Adaptive Cards input interface
  get value() {
    return this._selectedValue > 0 ? String(this._selectedValue) : undefined;
  }

  get isSet() {
    return this._selectedValue > 0;
  }

  isValid() {
    // No es required por defecto, así que siempre válido
    return true;
  }

  // Lo que necesita el SDK para serializar el valor al hacer Submit
  getJsonTypeName() {
    return "Input.Rating";
  }
}

// Registrar el custom element
AdaptiveCards.GlobalRegistry.elements.register(
  "Input.Rating",
  InputRating
);


// ============ 2) Conectar markdown-it con adaptivecards ============

// Instancia de markdown-it compartida
let md = null;

if (typeof markdownit !== "undefined") {
  md = markdownit({
    html: false,        // no permitir HTML raw por seguridad
    linkify: true,      // auto-detecta URLs y las hace clicables
    typographer: true,
    breaks: true         // saltos de línea = <br>
  });

  // Integración con la librería adaptivecards
  AdaptiveCards.AdaptiveCard.onProcessMarkdown = function (text, result) {
    result.outputHtml = md.render(text);
    result.didProcess = true;
  };
}

// ============ 3) Función exportada para burbujas normales del chat ============

/**
 * Convierte Markdown a HTML seguro para usar en burbujas del chat.
 * Si markdown-it no está cargado, devuelve el texto escapado con <br> en saltos de línea.
 */
function renderMarkdown(text) {
  if (!text) return '';
  if (md) {
    return md.render(text);
  }
  // Fallback sin markdown-it: escapa HTML y convierte \n a <br>
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');
}