
let codigoActual = {}
let articulos = JSON.parse(localStorage.getItem("articulosGuardados") || "[]")
  .map(a => ({ ...a, contenidoHTML: a.contenidoHTML ?? null }))
let materiasOrden = JSON.parse(localStorage.getItem("materiasOrden") || "{}")
let carpetas = JSON.parse(localStorage.getItem("carpetasMaterias") || "[]")
let normativaSeleccionada = null
let materiaSeleccionada = null
let resultadosBusqueda = []
let indiceResultado = 0
let ultimoTerminoBusqueda = ""
let documentosCargados = []
const elementosMarcados = new Set()
const buscadorInput = document.getElementById("buscadorInput")
const appRoot = document.getElementById("appRoot")
const documentoInput = document.getElementById("documentoInput")
const listaDocumentos = document.getElementById("listaDocumentos")
const visorDocumentos = document.getElementById("visorDocumentos")
const modalCarpeta = document.getElementById("modalCarpeta")
const modalCarpetaTitulo = document.getElementById("modalCarpetaTitulo")
const inputNombreCarpeta = document.getElementById("inputNombreCarpeta")
const modalCarpetaGuardar = document.getElementById("modalCarpetaGuardar")

function escaparComoHTML(texto) {
  if (texto === undefined || texto === null) return ""
  return String(texto)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
}

if (typeof pdfjsLib !== "undefined") {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js"
}

const ZOOM_STEP = 0.05
const MIN_ZOOM = 1.00
const MAX_ZOOM = 1.25
let zoomActual = 1
let pinchStartDistance = null
let pinchStartZoom = 1
let articuloArrastradoId = null
let materiaArrastrada = null
let materiaArrastradaNormativa = null
let materiaArrastradaCarpetaId = null
let carpetaEnEdicionId = null

documentoInput?.addEventListener("change", e => {
  const archivo = e.target.files?.[0]
  if (archivo) {
    procesarDocumento(archivo)
  }
  e.target.value = ""
})

const botonDocumentos = document.querySelector(".documentos-btn")

if (botonDocumentos) {
  ;["dragenter", "dragover"].forEach(evento => {
    botonDocumentos.addEventListener(evento, e => {
      e.preventDefault()
      e.stopPropagation()
      botonDocumentos.classList.add("arrastrando")
    })
  })

  ;["dragleave", "drop"].forEach(evento => {
    botonDocumentos.addEventListener(evento, e => {
      e.preventDefault()
      e.stopPropagation()
      botonDocumentos.classList.remove("arrastrando")
    })
  })

  botonDocumentos.addEventListener("drop", e => {
    const archivo = e.dataTransfer?.files?.[0]
    if (archivo) procesarDocumento(archivo)
  })
}

modalCarpetaGuardar?.addEventListener("click", confirmarCarpeta)
modalCarpeta?.addEventListener("click", e => {
  if (e.target === modalCarpeta) cerrarModalCarpeta()
})
inputNombreCarpeta?.addEventListener("keydown", e => {
  if (e.key === "Enter") {
    e.preventDefault()
    confirmarCarpeta()
  }
})

function obtenerExtension(nombre = "") {
  const partes = nombre.split(".")
  return partes.length > 1 ? partes.pop().toLowerCase() : ""
}

async function procesarDocumento(archivo) {
  const extension = obtenerExtension(archivo.name)
  const id = `${Date.now()}-${Math.random().toString(16).slice(2)}`
  const base = {
    id,
    nombre: archivo.name,
    extension,
    url: "",
    texto: "",
    mensaje: ""
  }

  try {
    if (extension === "pdf") {
      base.url = URL.createObjectURL(archivo)
      base.texto = await extraerTextoPdf(archivo)
    } else if (extension === "docx") {
      base.texto = await extraerTextoDocx(archivo)
    } else if (extension === "doc") {
      base.mensaje = "Vista previa limitada: descárgalo para abrirlo."
      base.url = URL.createObjectURL(archivo)
    } else {
      base.mensaje = "Formato no soportado para vista previa."
      base.url = URL.createObjectURL(archivo)
    }

    documentosCargados = [base, ...documentosCargados.filter(d => d.nombre !== base.nombre)]
    renderDocumentos()
    mostrarDocumento(base.id)
  } catch (err) {
    console.error("No se pudo procesar el documento", err)
    base.mensaje = "No se pudo leer el documento."
    documentosCargados = [base, ...documentosCargados]
    renderDocumentos()
    mostrarDocumento(base.id)
  }
}

async function extraerTextoDocx(archivo) {
  if (typeof JSZip === "undefined") {
    throw new Error("JSZip no está disponible para leer .docx")
  }

  const buffer = await archivo.arrayBuffer()
  const zip = await JSZip.loadAsync(buffer)
  const documento = zip.file("word/document.xml")

  if (!documento) {
    throw new Error("El archivo no contiene texto legible")
  }

  const xml = await documento.async("text")
  const parser = new DOMParser()
  const dom = parser.parseFromString(xml, "application/xml")
  const parrafos = Array.from(dom.getElementsByTagName("w:p"))

  const texto = parrafos
    .map(p => {
      const runs = p.getElementsByTagName("w:t")
      let acumulado = ""
      for (let i = 0; i < runs.length; i++) {
        acumulado += runs[i].textContent
      }
      return acumulado.trim()
    })
    .filter(Boolean)
    .join("\n\n")

  return texto || "No se encontró texto en el documento"
}

async function extraerTextoPdf(archivo) {
  if (typeof pdfjsLib === "undefined") {
    throw new Error("PDF.js no está disponible para leer el PDF")
  }

  const buffer = await archivo.arrayBuffer()
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise
  let texto = ""

  for (let i = 1; i <= pdf.numPages; i++) {
    const pagina = await pdf.getPage(i)
    const contenido = await pagina.getTextContent()
    const linea = contenido.items
      .map(item => (item.str || "").trim())
      .filter(Boolean)
      .join(" ")

    if (linea) {
      texto += linea + (i < pdf.numPages ? "\n\n" : "")
    }
  }

  return texto.trim() || "No se encontró texto en el PDF"
}

function renderDocumentos() {
  if (!listaDocumentos) return

  listaDocumentos.innerHTML = ""

  if (!documentosCargados.length) {
    const vacio = document.createElement("div")
    vacio.className = "documentos-vacio"
    vacio.textContent = "Aún no hay documentos cargados."
    listaDocumentos.appendChild(vacio)
    return
  }

  documentosCargados.forEach(doc => {
    const item = document.createElement("div")
    item.className = "documento-item"

    const info = document.createElement("div")
    info.className = "documento-info"

    const nombre = document.createElement("strong")
    nombre.textContent = doc.nombre

    const tipo = document.createElement("div")
    tipo.innerHTML = `<small>${doc.extension ? doc.extension.toUpperCase() : "Archivo"}</small>`

    info.appendChild(nombre)
    info.appendChild(tipo)

    const acciones = document.createElement("div")
    acciones.className = "documento-acciones"

    const ver = document.createElement("button")
    ver.className = "documento-ver"
    ver.type = "button"
    ver.textContent = "Ver"
    ver.addEventListener("click", () => mostrarDocumento(doc.id))

    const eliminar = document.createElement("button")
    eliminar.className = "documento-eliminar"
    eliminar.type = "button"
    eliminar.textContent = "Eliminar"
    eliminar.addEventListener("click", () => eliminarDocumento(doc.id))

    acciones.appendChild(ver)
    acciones.appendChild(eliminar)

    item.appendChild(info)
    item.appendChild(acciones)
    listaDocumentos.appendChild(item)
  })
}

function eliminarDocumento(id) {
  const doc = documentosCargados.find(d => d.id === id)
  documentosCargados = documentosCargados.filter(d => d.id !== id)

  if (doc?.url) {
    try { URL.revokeObjectURL(doc.url) } catch (e) { /* noop */ }
  }

  renderDocumentos()

  const vistaActual = visorDocumentos?.querySelector("h4")?.textContent
  if (doc && vistaActual === doc.nombre) {
    mostrarDocumento(documentosCargados[0]?.id)
  }
}

function mostrarDocumento(id, terminoBusqueda = "", indiceCoincidencia = null) {
  if (!visorDocumentos) return

  const doc = documentosCargados.find(d => d.id === id)
  visorDocumentos.innerHTML = ""
  visorDocumentos.dataset.docActual = doc?.id || ""

  const titulo = document.createElement("h4")
  titulo.textContent = doc ? doc.nombre : "Vista previa"
  visorDocumentos.appendChild(titulo)

  if (!doc) {
    const vacio = document.createElement("div")
    vacio.className = "documentos-vacio"
    vacio.textContent = "Selecciona un documento para verlo aquí."
    visorDocumentos.appendChild(vacio)
    return
  }

  if ((doc.extension === "pdf" || doc.extension === "docx") && doc.texto) {
    const texto = document.createElement("div")
    texto.className = "documento-texto"
    aplicarResaltadoEnTexto(doc.texto, texto, terminoBusqueda, indiceCoincidencia)
    visorDocumentos.appendChild(texto)
  } else if (doc.extension === "pdf" && doc.url) {
    const iframe = document.createElement("iframe")
    iframe.className = "documento-iframe"
    iframe.src = doc.url
    iframe.title = `Vista previa de ${doc.nombre}`
    visorDocumentos.appendChild(iframe)
  } else {
    const alerta = document.createElement("div")
    alerta.className = "documento-alerta"
    alerta.textContent = doc.mensaje || "No se pudo generar vista previa."
    visorDocumentos.appendChild(alerta)
  }

  if (doc.url) {
    const descarga = document.createElement("a")
    descarga.href = doc.url
    descarga.download = doc.nombre
    descarga.textContent = "Descargar original"
    descarga.style.fontWeight = "700"
    descarga.style.color = "var(--accent)"
    descarga.style.textDecoration = "none"
    descarga.style.marginTop = "6px"
    visorDocumentos.appendChild(descarga)
  }
}

function aplicarResaltadoEnTexto(textoFuente, contenedor, termino, indiceActivo) {
  const terminoLimpio = (termino || "").trim().toLowerCase()

  if (!terminoLimpio) {
    contenedor.textContent = textoFuente
    return
  }

  const textoPlano = textoFuente || ""
  const textoMin = textoPlano.toLowerCase()
  let pos = textoMin.indexOf(terminoLimpio)
  let ultimoCorte = 0
  let contador = 0

  if (pos === -1) {
    contenedor.textContent = textoPlano
    return
  }

  while (pos !== -1) {
    if (pos > ultimoCorte) {
      contenedor.appendChild(document.createTextNode(textoPlano.slice(ultimoCorte, pos)))
    }

    const marca = document.createElement("span")
    marca.className = "resaltado-busqueda-marca"
    marca.textContent = textoPlano.slice(pos, pos + terminoLimpio.length)
    if (indiceActivo !== null && contador === indiceActivo) {
      marca.classList.add("resaltado-busqueda-activo")
    }
    contenedor.appendChild(marca)

    ultimoCorte = pos + terminoLimpio.length
    pos = textoMin.indexOf(terminoLimpio, ultimoCorte)
    contador++
  }

  if (ultimoCorte < textoPlano.length) {
    contenedor.appendChild(document.createTextNode(textoPlano.slice(ultimoCorte)))
  }
}

function valorOrdenArticulo(a) {
  if (typeof a.orden === "number") return a.orden
  if (typeof a.numero === "number") return a.numero
  return 0
}

function guardarOrdenMaterias() {
  localStorage.setItem("materiasOrden", JSON.stringify(materiasOrden))
}

function siguienteOrdenMateria(normativa) {
  const ordenes = materiasOrden[normativa]
    ? Object.values(materiasOrden[normativa]).filter(n => typeof n === "number")
    : []

  if (!ordenes.length) return 1
  return Math.max(...ordenes) + 1
}

function asegurarOrdenMateria(normativa, materia) {
  if (!materiasOrden[normativa]) materiasOrden[normativa] = {}
  if (typeof materiasOrden[normativa][materia] !== "number") {
    materiasOrden[normativa][materia] = siguienteOrdenMateria(normativa)
    guardarOrdenMaterias()
  }
}

function ordenarMaterias(nombres, normativa) {
  nombres.forEach(nombre => asegurarOrdenMateria(normativa, nombre))
  return [...nombres].sort(
    (a, b) => (materiasOrden[normativa][a] ?? 0) - (materiasOrden[normativa][b] ?? 0)
  )
}

function guardarCarpetas() {
  localStorage.setItem("carpetasMaterias", JSON.stringify(carpetas))
}

function materiaEnCarpeta(normativa, materia) {
  const carpeta = carpetas.find(c =>
    c.materias?.some(m => m.materia === materia && m.normativa === normativa)
  )
  return carpeta ? carpeta.id : null
}

function removerMateriaDeCarpetas(normativa, materia) {
  let cambio = false
  carpetas = carpetas.map(carpeta => {
    const materias = (carpeta.materias || []).filter(
      m => !(m.materia === materia && m.normativa === normativa)
    )
    if (materias.length !== (carpeta.materias || []).length) cambio = true
    return { ...carpeta, materias }
  })

  if (cambio) guardarCarpetas()
  return cambio
}

function renombrarMateriaEnCarpetas(normativa, materiaAnterior, materiaNueva) {
  let cambio = false

  carpetas = carpetas.map(carpeta => {
    const materias = (carpeta.materias || []).map(m => {
      if (m.normativa === normativa && m.materia === materiaAnterior) {
        cambio = true
        return { ...m, materia: materiaNueva }
      }
      return m
    })

    const set = new Set()
    const materiasUnicas = materias.filter(m => {
      const key = `${m.normativa}||${m.materia}`
      if (set.has(key)) {
        cambio = true
        return false
      }
      set.add(key)
      return true
    })

    return { ...carpeta, materias: materiasUnicas }
  })

  if (cambio) guardarCarpetas()
}

function moverMateriaACarpeta(normativa, materia, carpetaId) {
  if (!carpetaId) return
  removerMateriaDeCarpetas(normativa, materia)
  carpetas = carpetas.map(carpeta => {
    if (carpeta.id !== carpetaId) return carpeta
    const materias = [...(carpeta.materias || [])]
    if (!materias.some(m => m.materia === materia && m.normativa === normativa)) {
      materias.push({ normativa, materia })
    }
    return { ...carpeta, materias }
  })
  guardarCarpetas()
  ordenarYMostrar()
}

function moverMateriaAFueraDeCarpeta(materia, normativa) {
  const cambio = removerMateriaDeCarpetas(normativa, materia)
  if (cambio) ordenarYMostrar()
}

function aplicarOrdenMateriasDesdeDOM(lista, normativa) {
  if (!lista) return
  const items = Array.from(lista.querySelectorAll(".sidebarItem"))
  if (!items.length) return

  if (!materiasOrden[normativa]) materiasOrden[normativa] = {}

  items.forEach((item, index) => {
    materiasOrden[normativa][item.dataset.materia] = index + 1
  })

  guardarOrdenMaterias()
}

function siguienteOrdenPara(normativa, materia) {
  const ordenes = articulos
    .filter(a => a.normativa === normativa && a.materia === materia)
    .map(valorOrdenArticulo)

  if (!ordenes.length) return 1
  return Math.max(...ordenes) + 1
}

function aplicarZoom(nivel) {
  const limitado = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, nivel))
  zoomActual = limitado
  document.documentElement.style.setProperty("--zoom-scale", limitado.toFixed(3))
}

function distanciaEntreToques(touches) {
  if (touches.length < 2) return 0
  const [a, b] = touches
  const dx = a.clientX - b.clientX
  const dy = a.clientY - b.clientY
  return Math.hypot(dx, dy)
}

function esCampoEditable(elem) {
  if (!elem) return null
  const destino = elem.closest(
    "input, textarea, [contenteditable=\"true\"], .articulo-contenido, .nota-box"
  )
  return destino || (elem.isContentEditable ? elem : null)
}

let menuContextual
let botonCopiarMenu
let botonPegarMenu
let objetivoContextual

function crearMenuContextual() {
  const menu = document.createElement("div")
  menu.className = "menu-contextual"

  const copiar = document.createElement("button")
  copiar.type = "button"
  copiar.textContent = "Copiar"

  const pegar = document.createElement("button")
  pegar.type = "button"
  pegar.textContent = "Pegar"

  menu.appendChild(copiar)
  menu.appendChild(pegar)

  return { menu, copiar, pegar }
}

function ocultarMenuContextual() {
  if (menuContextual) menuContextual.style.display = "none"
  objetivoContextual = null
}

function seleccionPerteneceAlObjetivo(objetivo, seleccion) {
  if (!seleccion) return false
  if (!seleccion.anchorNode) return false
  if (seleccion.anchorNode === objetivo) return true
  return objetivo.contains(seleccion.anchorNode)
}

function textoSeleccionadoEnObjetivo(objetivo) {
  if (!objetivo) return ""

  if (objetivo instanceof HTMLInputElement || objetivo instanceof HTMLTextAreaElement) {
    const inicio = objetivo.selectionStart ?? 0
    const fin = objetivo.selectionEnd ?? 0
    if (inicio !== fin) return objetivo.value.slice(inicio, fin)
    return objetivo.value
  }

  const seleccion = window.getSelection()
  if (seleccion && seleccion.toString().trim() && seleccionPerteneceAlObjetivo(objetivo, seleccion)) {
    return seleccion.toString()
  }

  return objetivo.textContent || ""
}

function actualizarEstadoMenuContextual() {
  if (!menuContextual) return

  const hayObjetivo = Boolean(objetivoContextual)
  const textoSeleccion = hayObjetivo ? textoSeleccionadoEnObjetivo(objetivoContextual).trim() : ""

  const soporteClipboard = Boolean(navigator.clipboard)
  const puedeCopiar = soporteClipboard && hayObjetivo && Boolean(textoSeleccion)
  const puedePegar = soporteClipboard && hayObjetivo && Boolean(navigator.clipboard.readText)

  botonCopiarMenu.disabled = !puedeCopiar
  botonPegarMenu.disabled = !puedePegar
}

function mostrarMenuContextual(x, y) {
  if (!menuContextual) return

  menuContextual.style.display = "flex"

  const ancho = menuContextual.offsetWidth
  const alto = menuContextual.offsetHeight
  const margen = 12

  let posX = x
  let posY = y

  const maxX = window.innerWidth - ancho - margen
  const maxY = window.innerHeight - alto - margen

  posX = Math.min(posX, Math.max(margen, maxX))
  posY = Math.min(posY, Math.max(margen, maxY))

  menuContextual.style.left = `${posX}px`
  menuContextual.style.top = `${posY}px`
}

function enfocarObjetivoParaPegar() {
  if (!objetivoContextual) return
  const seleccion = window.getSelection()

  if (
    objetivoContextual instanceof HTMLInputElement ||
    objetivoContextual instanceof HTMLTextAreaElement
  ) {
    objetivoContextual.focus({ preventScroll: true })
    return
  }

  if (seleccion && seleccionPerteneceAlObjetivo(objetivoContextual, seleccion)) return

  const rango = document.createRange()
  const ultimoNodo = objetivoContextual.lastChild
  if (ultimoNodo) {
    rango.setStartAfter(ultimoNodo)
  } else {
    rango.setStart(objetivoContextual, 0)
  }
  rango.collapse(true)

  if (seleccion) {
    seleccion.removeAllRanges()
    seleccion.addRange(rango)
  }

  objetivoContextual.focus({ preventScroll: true })
}

function pegarEnObjetivo(texto) {
  if (!objetivoContextual || !texto) return

  if (objetivoContextual instanceof HTMLInputElement || objetivoContextual instanceof HTMLTextAreaElement) {
    const inicio = objetivoContextual.selectionStart ?? objetivoContextual.value.length
    const fin = objetivoContextual.selectionEnd ?? objetivoContextual.value.length
    try {
      objetivoContextual.setRangeText(texto, inicio, fin, "end")
    } catch (error) {
      const nuevoValor =
        objetivoContextual.value.slice(0, inicio) + texto + objetivoContextual.value.slice(fin)

      objetivoContextual.value = nuevoValor

      const nuevaPos = inicio + texto.length
      objetivoContextual.setSelectionRange(nuevaPos, nuevaPos)
    }

    objetivoContextual.dispatchEvent(new Event("input", { bubbles: true }))
    return
  }

  enfocarObjetivoParaPegar()

  const seleccion = window.getSelection()
  if (seleccion && seleccion.rangeCount > 0) {
    const rango = seleccion.getRangeAt(0)
    const rangoClonado = rango.cloneRange()

    const pudoInsertar =
      typeof document.queryCommandSupported === "function" &&
      document.queryCommandSupported("insertText")
        ? document.execCommand("insertText", false, texto)
        : false

    if (!pudoInsertar) {
      rangoClonado.deleteContents()
      rangoClonado.insertNode(document.createTextNode(texto))
      rangoClonado.collapse(false)
      seleccion.removeAllRanges()
      seleccion.addRange(rangoClonado)
    }
  } else {
    objetivoContextual.textContent += texto
  }

  objetivoContextual.dispatchEvent(new Event("input", { bubbles: true }))
}

function inicializarMenuContextual() {
  const { menu, copiar, pegar } = crearMenuContextual()
  menuContextual = menu
  botonCopiarMenu = copiar
  botonPegarMenu = pegar

  document.body.appendChild(menuContextual)

  document.addEventListener("contextmenu", e => {
    const editable = esCampoEditable(e.target)

    if (!editable) {
      ocultarMenuContextual()
      return
    }

    e.preventDefault()
    objetivoContextual = editable
    actualizarEstadoMenuContextual()
    mostrarMenuContextual(e.clientX, e.clientY)
  })

  botonCopiarMenu.addEventListener("click", async () => {
    if (!objetivoContextual || !navigator.clipboard?.writeText) return

    const texto = textoSeleccionadoEnObjetivo(objetivoContextual)
    if (!texto.trim()) return

    try {
      await navigator.clipboard.writeText(texto)
    } catch (error) {
      console.error("No se pudo copiar", error)
    }

    ocultarMenuContextual()
  })

  botonPegarMenu.addEventListener("click", async () => {
    if (!objetivoContextual || !navigator.clipboard?.readText) return

    try {
      const texto = await navigator.clipboard.readText()
      pegarEnObjetivo(texto)
    } catch (error) {
      console.error("No se pudo pegar", error)
    }

    ocultarMenuContextual()
  })

  document.addEventListener(
    "pointerdown",
    e => {
      if (e.button !== 0) return
      if (!menuContextual || menuContextual.style.display === "none") return
      if (!menuContextual.contains(e.target)) ocultarMenuContextual()
    },
    true
  )

  document.addEventListener("click", e => {
    if (!menuContextual || menuContextual.style.display === "none") return
    if (!menuContextual.contains(e.target)) ocultarMenuContextual()
  })

  window.addEventListener("resize", ocultarMenuContextual)
  window.addEventListener("scroll", ocultarMenuContextual, true)
}

function aplicarOrdenDesdeDOM(contenedorLista, normativa, materia) {
  if (!contenedorLista) return

  const idsEnOrden = Array.from(contenedorLista.querySelectorAll(".articulo"))
    .map(el => Number(el.dataset.id))
    .filter(Boolean)

  idsEnOrden.forEach((id, index) => {
    const art = articulos.find(a => a.id === id)
    if (art && art.normativa === normativa && art.materia === materia) {
      art.orden = index + 1
    }
  })

  guardarLocal()
  ordenarYMostrar()
}

function activarArrastreMateria(item, lista, normativa) {
  item.draggable = true

  item.addEventListener("dragstart", e => {
    materiaArrastrada = item.dataset.materia
    materiaArrastradaNormativa = normativa
    materiaArrastradaCarpetaId = item.dataset.carpetaId || null
    item.classList.add("arrastrando")
    if (e.dataTransfer) {
      e.dataTransfer.effectAllowed = "move"
      e.dataTransfer.setData("text/plain", materiaArrastrada)
    }
  })

  item.addEventListener("dragover", e => {
    e.preventDefault()
    if (e.dataTransfer) e.dataTransfer.dropEffect = "move"
    if (!materiaArrastrada || item.dataset.materia === materiaArrastrada) return

    const rect = item.getBoundingClientRect()
    const margenCambio = Math.max(10, rect.height * 0.35)
    const limiteArriba = rect.top + margenCambio
    const limiteAbajo = rect.bottom - margenCambio
    const posicionY = e.clientY

    let moverDespues = null
    if (posicionY <= limiteArriba) {
      moverDespues = false
    } else if (posicionY >= limiteAbajo) {
      moverDespues = true
    } else {
      return
    }

    const arrastrandoElem = lista.querySelector(
      `.sidebarItem[data-materia="${materiaArrastrada}"][data-normativa="${normativa}"]`
    )

    if (!arrastrandoElem || arrastrandoElem === item) return

    if (moverDespues) {
      lista.insertBefore(arrastrandoElem, item.nextSibling)
    } else {
      lista.insertBefore(arrastrandoElem, item)
    }
  })

  item.addEventListener("drop", e => e.preventDefault())

  item.addEventListener("dragend", () => {
    item.classList.remove("arrastrando")
    aplicarOrdenMateriasDesdeDOM(lista, normativa)
    materiaArrastrada = null
    materiaArrastradaNormativa = null
    materiaArrastradaCarpetaId = null
  })
}

function activarArrastreArticulo(box, normativa, materia) {
  box.draggable = true

  const bloquearSiEditable = e => {
    if (esCampoEditable(e.target)) {
      box.draggable = false
    } else {
      box.draggable = true
    }
  }

  const restaurarArrastre = () => {
    if (!box.draggable) box.draggable = true
  }

  box.addEventListener("mousedown", bloquearSiEditable)
  box.addEventListener("touchstart", bloquearSiEditable)
  box.addEventListener("mouseup", restaurarArrastre)
  box.addEventListener("mouseleave", restaurarArrastre)
  box.addEventListener("touchend", restaurarArrastre)
  box.addEventListener("touchcancel", restaurarArrastre)

  box.addEventListener("dragstart", e => {
    if (esCampoEditable(e.target)) {
      box.draggable = false
      setTimeout(() => {
        box.draggable = true
      }, 0)
      e.preventDefault()
      return
    }

    articuloArrastradoId = box.dataset.id
    box.classList.add("arrastrando")
    if (e.dataTransfer) {
      e.dataTransfer.effectAllowed = "move"
      e.dataTransfer.setData("text/plain", box.dataset.id)
    }
  })

  box.addEventListener("dragenter", () => {
    if (box.dataset.id === articuloArrastradoId) return
    box.classList.add("drag-over")
  })

  box.addEventListener("dragleave", () => {
    box.classList.remove("drag-over")
  })

  box.addEventListener("dragover", e => {
    e.preventDefault()
    if (e.dataTransfer) e.dataTransfer.dropEffect = "move"

    const objetivo = e.currentTarget
    if (!articuloArrastradoId || objetivo.dataset.id === articuloArrastradoId) return

    const rect = objetivo.getBoundingClientRect()
    const margenCambio = Math.max(12, rect.height * 0.28)
    const antesLimite = rect.top + margenCambio
    const despuesLimite = rect.bottom - margenCambio
    const posicionY = e.clientY

    // Solo reordenar cuando el puntero se acerca a los bordes para que
    // el cambio ocurra antes y con menos fricción.
    let desplazarDespues = null
    if (posicionY <= antesLimite) {
      desplazarDespues = false
    } else if (posicionY >= despuesLimite) {
      desplazarDespues = true
    } else {
      return
    }

    const contenedor = objetivo.parentElement
    const arrastrandoElem = contenedor.querySelector(
      `.articulo[data-id="${articuloArrastradoId}"]`
    )

    if (!arrastrandoElem || arrastrandoElem === objetivo) return

    if (desplazarDespues) {
      contenedor.insertBefore(arrastrandoElem, objetivo.nextSibling)
    } else {
      contenedor.insertBefore(arrastrandoElem, objetivo)
    }
  })

  box.addEventListener("drop", e => {
    e.preventDefault()
    box.classList.remove("drag-over")
  })

  box.addEventListener("dragend", () => {
    box.classList.remove("arrastrando")
    box.classList.remove("drag-over")
    const contenedor = box.parentElement
    aplicarOrdenDesdeDOM(contenedor, normativa, materia)
    articuloArrastradoId = null
  })
}

document.addEventListener(
  "wheel",
  e => {
    if (!e.ctrlKey) return
    e.preventDefault()
    const delta = e.deltaY > 0 ? -ZOOM_STEP : ZOOM_STEP
    aplicarZoom(zoomActual + delta)
  },
  { passive: false }
)

document.addEventListener("keydown", e => {
  const ctrl = e.ctrlKey || e.metaKey
  if (!ctrl) return

  if (e.key === "+" || e.key === "=") {
    e.preventDefault()
    aplicarZoom(zoomActual + ZOOM_STEP)
  } else if (e.key === "-" || e.key === "_") {
    e.preventDefault()
    aplicarZoom(zoomActual - ZOOM_STEP)
  } else if (e.key === "0") {
    e.preventDefault()
    aplicarZoom(1)
  }
})

if (appRoot) {
  appRoot.addEventListener("touchstart", e => {
    if (e.touches.length === 2) {
      pinchStartDistance = distanciaEntreToques(e.touches)
      pinchStartZoom = zoomActual
    }
  })

  appRoot.addEventListener(
    "touchmove",
    e => {
      if (e.touches.length === 2 && pinchStartDistance) {
        e.preventDefault()
        const nuevaDistancia = distanciaEntreToques(e.touches)
        const factor = pinchStartDistance ? nuevaDistancia / pinchStartDistance : 1
        aplicarZoom(pinchStartZoom * factor)
      }
    },
    { passive: false }
  )

  const limpiarPinch = e => {
    if (e.touches && e.touches.length >= 2) return
    pinchStartDistance = null
    pinchStartZoom = zoomActual
  }

  appRoot.addEventListener("touchend", limpiarPinch)
  appRoot.addEventListener("touchcancel", limpiarPinch)
}

aplicarZoom(zoomActual)

function abrirBuscador() {
  const panel = document.getElementById("buscadorFlotante")
  panel.classList.add("activo")

  buscadorInput.focus()
  buscadorInput.select()

  if (ultimoTerminoBusqueda.trim() !== "") {
    buscadorInput.value = ultimoTerminoBusqueda
    ejecutarBusqueda(ultimoTerminoBusqueda)
  }
}

function cerrarBuscador() {
  document.getElementById("buscadorFlotante").classList.remove("activo")
  limpiarResaltadosBusqueda()
}

function limpiarResaltadosBusqueda() {
  resultadosBusqueda = []
  indiceResultado = 0

  elementosMarcados.forEach(el => {
    if (el.dataset.busquedaOriginalHtml !== undefined) {
      el.innerHTML = el.dataset.busquedaOriginalHtml
    } else if (el.isContentEditable) {
      const plano = el.textContent
      el.textContent = plano
    }
    if (el.normalize) el.normalize()
    delete el.dataset.busquedaOriginalHtml
  })
  elementosMarcados.clear()

  document.querySelectorAll(".articulo").forEach(a => {
    a.classList.remove("resaltado-busqueda")
    a.classList.remove("resaltado-busqueda-activo")
  })

  document.querySelectorAll(".resaltado-busqueda-activo").forEach(el =>
    el.classList.remove("resaltado-busqueda-activo")
  )

  if (visorDocumentos?.dataset.docActual) {
    const doc = documentosCargados.find(d => d.id === visorDocumentos.dataset.docActual)
    if (doc) mostrarDocumento(doc.id)
  }
}

function resaltarResultadoActual(mantenerFocoBuscador = false) {
  const habiaFocoEnBuscador = mantenerFocoBuscador || document.activeElement === buscadorInput

  document.querySelectorAll(".resaltado-busqueda-activo").forEach(el =>
    el.classList.remove("resaltado-busqueda-activo")
  )

  document.querySelectorAll(".resaltado-busqueda-marca-activa").forEach(el =>
    el.classList.remove("resaltado-busqueda-marca-activa")
  )

  if (!resultadosBusqueda.length) return

  const actual = resultadosBusqueda[indiceResultado]
  const articulo = actual.articulo

  if (actual.tipo === "documento") {
    const doc = documentosCargados.find(d => d.id === actual.documentoId)
    if (doc) {
      mostrarDocumento(doc.id, ultimoTerminoBusqueda, actual.indiceDocumento)
      const marcaActiva = visorDocumentos?.querySelector(".resaltado-busqueda-activo")
      if (marcaActiva) {
        marcaActiva.scrollIntoView({ behavior: "smooth", block: "center" })
      } else {
        visorDocumentos?.scrollIntoView({ behavior: "smooth", block: "center" })
      }
    }
    return
  }

  if (articulo && !articulo.classList.contains("abierto")) {
    articulo.classList.add("abierto")

    const nota = articulo.querySelector(".nota-box")
    const resizer = articulo.querySelector(".nota-resizer")
    const botonNota = articulo.querySelector(".btn-note")
    const botonEliminarNota = articulo.querySelector(".btn-note-delete")

    if (botonNota) botonNota.style.display = "inline-block"
    if (nota && nota.value.trim() !== "") {
      nota.style.display = "block"
      if (resizer) resizer.style.display = "block"
      if (botonEliminarNota) botonEliminarNota.style.display = "inline-block"
    }
  }

  if (articulo) articulo.classList.add("resaltado-busqueda")

  if (actual.tipo === "nota") {
    actual.nodo.classList.add("resaltado-busqueda-activo")
    actual.nodo.setSelectionRange(actual.inicio, actual.fin)
    actual.nodo.scrollIntoView({ behavior: "smooth", block: "center" })

    if (habiaFocoEnBuscador) {
      buscadorInput.focus({ preventScroll: true })
      const caret = buscadorInput.value.length
      buscadorInput.setSelectionRange(caret, caret)
    }
    return
  }

  if (actual.tipo === "marca" && actual.nodo) {
    actual.nodo.classList.add("resaltado-busqueda-marca-activa")
    if (articulo) articulo.classList.add("resaltado-busqueda-activo")
    const destino = actual.nodo.closest(".articulo") || actual.nodo
    destino.scrollIntoView({ behavior: "smooth", block: "center" })
  }

  if (habiaFocoEnBuscador) {
    buscadorInput.focus({ preventScroll: true })
    const caret = buscadorInput.value.length
    buscadorInput.setSelectionRange(caret, caret)
  }
}

function ejecutarBusqueda(termino) {
  ultimoTerminoBusqueda = termino
  const limpio = termino.trim().toLowerCase()
  const info = document.getElementById("buscadorInfo")

  limpiarResaltadosBusqueda()
  resultadosBusqueda = []
  indiceResultado = 0

  if (!limpio) {
    info.textContent = "Escribe una palabra clave"
    return
  }

  document.querySelectorAll(".articulo").forEach(art => {
    const titulo = art.querySelector(".articulo-titulo-texto") || art.querySelector(".articulo-titulo")
    const contenido = art.querySelector(".articulo-contenido")
    const nota = art.querySelector(".nota-box")

    marcarCoincidencias(titulo, limpio, art)
    marcarCoincidencias(contenido, limpio, art)
    registrarCoincidenciasNota(nota, limpio, art)
  })

  documentosCargados.forEach(doc => registrarCoincidenciasDocumento(doc, limpio))

  if (!resultadosBusqueda.length) {
    info.textContent = "Sin resultados"
    return
  }

  info.textContent = `${resultadosBusqueda.length} resultado(s)`
  resaltarResultadoActual(true)
}

function marcarCoincidencias(elemento, termino, articulo) {
  if (!elemento) return

  const largo = termino.length
  if (!largo) return

  if (elemento.dataset.busquedaOriginalHtml === undefined) {
    elemento.dataset.busquedaOriginalHtml = elemento.innerHTML
    elementosMarcados.add(elemento)
  }

  const nodosTexto = []
  const walker = document.createTreeWalker(elemento, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      if (!node.textContent.trim()) return NodeFilter.FILTER_REJECT
      if (node.parentElement && node.parentElement.closest(".resaltado-busqueda-marca")) {
        return NodeFilter.FILTER_REJECT
      }
      return NodeFilter.FILTER_ACCEPT
    }
  })

  let nodo = walker.nextNode()
  while (nodo) {
    nodosTexto.push(nodo)
    nodo = walker.nextNode()
  }

  let hayCoincidencias = false

  nodosTexto.forEach(textNode => {
    const textoOriginal = textNode.textContent
    const textoMin = textoOriginal.toLowerCase()
    let indice = textoMin.indexOf(termino)
    if (indice === -1) return

    hayCoincidencias = true
    const frag = document.createDocumentFragment()
    let ultimoCorte = 0

    while (indice !== -1) {
      if (indice > ultimoCorte) {
        frag.appendChild(document.createTextNode(textoOriginal.slice(ultimoCorte, indice)))
      }

      const marca = document.createElement("span")
      marca.className = "resaltado-busqueda-marca"
      marca.textContent = textoOriginal.slice(indice, indice + largo)
      resultadosBusqueda.push({ tipo: "marca", nodo: marca, articulo })
      frag.appendChild(marca)

      ultimoCorte = indice + largo
      indice = textoMin.indexOf(termino, ultimoCorte)
    }

    if (ultimoCorte < textoOriginal.length) {
      frag.appendChild(document.createTextNode(textoOriginal.slice(ultimoCorte)))
    }

    if (textNode.parentNode) {
      textNode.parentNode.replaceChild(frag, textNode)
    }
  })

  if (hayCoincidencias) articulo.classList.add("resaltado-busqueda")
}

function registrarCoincidenciasNota(nota, termino, articulo) {
  if (!nota) return

  const texto = nota.value.toLowerCase()
  const largo = termino.length
  if (!largo) return

  let pos = texto.indexOf(termino)
  let hayCoincidencias = false

  while (pos !== -1) {
    resultadosBusqueda.push({
      tipo: "nota",
      nodo: nota,
      articulo,
      inicio: pos,
      fin: pos + largo
    })
    hayCoincidencias = true
    pos = texto.indexOf(termino, pos + largo)
  }

  if (hayCoincidencias) articulo.classList.add("resaltado-busqueda")
}

function registrarCoincidenciasDocumento(doc, termino) {
  if (!doc?.texto) return

  const textoMin = doc.texto.toLowerCase()
  const largo = termino.length
  if (!largo) return

  let pos = textoMin.indexOf(termino)
  let coincidencia = 0

  while (pos !== -1) {
    resultadosBusqueda.push({
      tipo: "documento",
      documentoId: doc.id,
      inicio: pos,
      fin: pos + largo,
      indiceDocumento: coincidencia
    })
    coincidencia++
    pos = textoMin.indexOf(termino, pos + largo)
  }
}

function irAlSiguienteResultado(mantenerFocoBuscador = false) {
  if (!resultadosBusqueda.length) return
  indiceResultado = (indiceResultado + 1) % resultadosBusqueda.length
  resaltarResultadoActual(mantenerFocoBuscador)
}

function reaplicarBusqueda() {
  const panel = document.getElementById("buscadorFlotante")
  if (panel.classList.contains("activo") && ultimoTerminoBusqueda.trim() !== "") {
    ejecutarBusqueda(ultimoTerminoBusqueda)
  }
}

window.addEventListener("keydown", e => {
  const isCtrlF = (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "f"
  const isF3 = e.key === "F3"
  if (isCtrlF || isF3) {
    e.preventDefault()
    abrirBuscador()
  }
})

buscadorInput.addEventListener("input", e => {
  ejecutarBusqueda(e.target.value)
})

buscadorInput.addEventListener("keydown", e => {
  if (e.key === "Enter") {
    e.preventDefault()
    irAlSiguienteResultado(true)
  }
  if (e.key === "Escape") {
    cerrarBuscador()
  }
})

function toggleModo() {
  document.body.classList.toggle("oscuro")
}

function cargarNormativa() {
  sincronizarEdiciones()

  const val = document.getElementById("normativa").value
  const archivo = val === "civil" ? "codigo_civil_pdf.json" : "codigo_penal_pdf.json"

  fetch(archivo)
    .then(r => r.json())
    .then(d => {
      codigoActual = d
      ordenarYMostrar()
    })
    .catch(() => mostrarError("No se pudo cargar el archivo"))
}

cargarNormativa()
document.getElementById("normativa").addEventListener("change", cargarNormativa)

function mostrarError(msg) {
  const box = document.getElementById("errorBox")
  box.textContent = msg
  box.style.display = "block"
}

function obtenerColorMateria(nombre) {
  const art = articulos.find(a => a.materia === nombre && a.color)
  return art ? art.color : "#1e3a8a"
}

function agregarArticulo() {
  const num = parseInt(document.getElementById("numeroArticulo").value)
  const mat = document.getElementById("materiaArticulo").value.trim()

  document.getElementById("errorBox").style.display = "none"

  if (!num || !mat) return mostrarError("Debe ingresar número y materia")

  const cont = codigoActual[num]
  if (!cont) return mostrarError("Ese artículo no existe en la base")

  const norm = document.getElementById("normativa").value

  asegurarOrdenMateria(norm, mat)

  articulos.push({
    id: Date.now(),
    numero: num,
    materia: mat,
    contenido: cont,
    contenidoEditado: null,
    contenidoHTML: null,
    normativa: norm,
    nota: "",
    color: obtenerColorMateria(mat),
    tituloPersonalizado: "",
    orden: siguienteOrdenPara(norm, mat)
  })

  normativaSeleccionada = norm
  materiaSeleccionada = mat

  guardarLocal()
  ordenarYMostrar()

  document.getElementById("numeroArticulo").value = ""
  document.getElementById("materiaArticulo").value = ""
  document.getElementById("numeroArticulo").focus()
}

function ordenarYMostrar() {
  sincronizarEdiciones()

  const sidebar = document.getElementById("sidebarMaterias")
  const contenedorArticulos = document.getElementById("contenidoArticulos")

  sidebar.innerHTML = ""
  contenedorArticulos.innerHTML = ""

  const accionesSidebar = document.createElement("div")
  accionesSidebar.className = "sidebarAcciones"
  const tituloSidebar = document.createElement("div")
  tituloSidebar.className = "sidebarAccionesTitulo"
  tituloSidebar.textContent = "Organiza tus materias"
  const nuevaCarpetaBtn = document.createElement("button")
  nuevaCarpetaBtn.className = "btn-carpeta"
  nuevaCarpetaBtn.type = "button"
  nuevaCarpetaBtn.textContent = "Nueva carpeta"
  nuevaCarpetaBtn.addEventListener("click", () => abrirModalCarpeta())
  accionesSidebar.appendChild(tituloSidebar)
  accionesSidebar.appendChild(nuevaCarpetaBtn)
  sidebar.appendChild(accionesSidebar)

  const seccionCarpetas = document.createElement("div")
  seccionCarpetas.className = "carpetasSection"
  const tituloCarpetas = document.createElement("div")
  tituloCarpetas.className = "sidebarGroupTitle"
  tituloCarpetas.textContent = "Carpetas"
  const listaCarpetas = document.createElement("div")
  listaCarpetas.className = "carpetasLista"
  seccionCarpetas.appendChild(tituloCarpetas)
  seccionCarpetas.appendChild(listaCarpetas)
  sidebar.appendChild(seccionCarpetas)

  const agrupado = { civil: {}, penal: {} }

  articulos.forEach(a => {
    if (!agrupado[a.normativa]) agrupado[a.normativa] = {}
    if (!agrupado[a.normativa][a.materia]) agrupado[a.normativa][a.materia] = []
    agrupado[a.normativa][a.materia].push(a)
  })

  const combos = []
  let tieneSeleccionValida = false

  carpetas.forEach(carpeta => {
    (carpeta.materias || []).forEach(({ normativa, materia }) => {
      if (!agrupado[normativa]?.[materia]) return
      if (!combos.some(c => c.normativa === normativa && c.materia === materia)) {
        combos.push({ normativa, materia, items: agrupado[normativa][materia] })
      }
      if (normativaSeleccionada === normativa && materiaSeleccionada === materia) {
        tieneSeleccionValida = true
      }
    })
  })

  renderizarCarpetasSidebar(listaCarpetas, agrupado, sidebar)

  const ordenNormativas = ["civil", "penal"]
  ordenNormativas.forEach(norm => {
    const materiasObj = agrupado[norm]
    const nombresMaterias = Object.keys(materiasObj || {})
    if (!nombresMaterias.length) return

    const grupo = document.createElement("div")
    grupo.className = "sidebarGroup"

    const tituloGrupo = document.createElement("div")
    tituloGrupo.className = "sidebarGroupTitle"
    tituloGrupo.textContent = norm === "civil" ? "Código Civil" : "Código Penal"
    grupo.appendChild(tituloGrupo)

    const listaMaterias = document.createElement("div")
    listaMaterias.className = "sidebarGroupList"

    const habilitarDropSalida = elemento => {
      elemento.addEventListener("dragover", e => {
        if (!materiaArrastrada) return
        e.preventDefault()
        elemento.classList.add("drop-activa")
        if (e.dataTransfer) e.dataTransfer.dropEffect = "move"
      })

      elemento.addEventListener("dragleave", () => {
        elemento.classList.remove("drop-activa")
      })

      elemento.addEventListener("drop", e => {
        if (!materiaArrastrada) return
        e.preventDefault()
        elemento.classList.remove("drop-activa")
        const normOrigen = materiaArrastradaNormativa || norm
        setTimeout(() => moverMateriaAFueraDeCarpeta(materiaArrastrada, normOrigen), 0)
      })
    }

    habilitarDropSalida(listaMaterias)
    grupo.appendChild(listaMaterias)

    const nombresOrdenados = ordenarMaterias(nombresMaterias, norm).filter(
      m => !materiaEnCarpeta(norm, m)
    )

    nombresOrdenados.forEach(m => {
      const item = document.createElement("div")
      item.className = "sidebarItem"
      item.textContent = m
      item.dataset.normativa = norm
      item.dataset.materia = m
      item.style.borderLeftColor = obtenerColorMateria(m)

      item.addEventListener("click", () => {
        normativaSeleccionada = norm
        materiaSeleccionada = m
        sidebar.querySelectorAll(".sidebarItem").forEach(i => i.classList.remove("activa"))
        item.classList.add("activa")
        mostrarArticulosDeMateria(norm, m, agrupado[norm][m])
      })

      if (norm === normativaSeleccionada && m === materiaSeleccionada) {
        item.classList.add("activa")
        tieneSeleccionValida = true
      }

      activarArrastreMateria(item, listaMaterias, norm)

      listaMaterias.appendChild(item)
      if (!combos.some(c => c.normativa === norm && c.materia === m)) {
        combos.push({ normativa: norm, materia: m })
      }
    })

    sidebar.appendChild(grupo)
  })

  if (!combos.length) {
    contenedorArticulos.innerHTML = `
      <div class="estado-vacio">
        <h2>No hay materias aún</h2>
        <p>Agrega artículos para crear automáticamente tu primera materia.</p>
      </div>
    `

    return
  }

  if (!tieneSeleccionValida) {
    normativaSeleccionada = combos[0].normativa
    materiaSeleccionada = combos[0].materia
  }

  const activo = sidebar.querySelector(
    `.sidebarItem[data-normativa="${normativaSeleccionada}"][data-materia="${materiaSeleccionada}"]`
  )
  if (activo) activo.classList.add("activa")

  const itemsSeleccionados =
    agrupado[normativaSeleccionada] && agrupado[normativaSeleccionada][materiaSeleccionada]
      ? agrupado[normativaSeleccionada][materiaSeleccionada]
      : []

  mostrarArticulosDeMateria(normativaSeleccionada, materiaSeleccionada, itemsSeleccionados)
  reaplicarBusqueda()
}

function renderizarCarpetasSidebar(contenedor, agrupado, sidebar) {
  contenedor.innerHTML = ""

  if (!carpetas.length) {
    const vacio = document.createElement("div")
    vacio.className = "carpetaVacia"
    vacio.textContent = "Crea carpetas y arrastra tus materias aquí"
    contenedor.appendChild(vacio)
    return
  }

  carpetas.forEach(carpeta => {
    const card = document.createElement("div")
    card.className = "carpetaBox"
    card.dataset.carpetaId = carpeta.id

    const header = document.createElement("div")
    header.className = "carpetaHeader"

    const nombreBtn = document.createElement("button")
    nombreBtn.className = "carpetaNombre"
    nombreBtn.type = "button"
    nombreBtn.textContent = carpeta.nombre || "Carpeta sin título"
    nombreBtn.addEventListener("click", () => abrirModalCarpeta(carpeta.id))

    const eliminarBtn = document.createElement("button")
    eliminarBtn.className = "carpetaEliminar"
    eliminarBtn.type = "button"
    eliminarBtn.textContent = "Borrar"
    eliminarBtn.addEventListener("click", () => eliminarCarpeta(carpeta.id))

    header.appendChild(nombreBtn)
    header.appendChild(eliminarBtn)
    card.appendChild(header)

    const lista = document.createElement("div")
    lista.className = "sidebarGroupList carpetaLista"
    lista.dataset.carpetaId = carpeta.id

    const materiasDisponibles = (carpeta.materias || []).filter(m =>
      agrupado[m.normativa]?.[m.materia]
    )

    lista.addEventListener("dragover", e => {
      if (!materiaArrastrada) return
      e.preventDefault()
      lista.classList.add("carpetaLista-drop")
      if (e.dataTransfer) e.dataTransfer.dropEffect = "move"
    })

    lista.addEventListener("dragleave", () => {
      lista.classList.remove("carpetaLista-drop")
    })

    lista.addEventListener("drop", e => {
      if (!materiaArrastrada || !materiaArrastradaNormativa) return
      e.preventDefault()
      lista.classList.remove("carpetaLista-drop")
      moverMateriaACarpeta(materiaArrastradaNormativa, materiaArrastrada, carpeta.id)
    })

    if (!materiasDisponibles.length) {
      const aviso = document.createElement("div")
      aviso.className = "carpetaVacia"
      aviso.textContent = "Arrastra materias aquí"
      lista.appendChild(aviso)
    } else {
      materiasDisponibles.forEach(({ normativa, materia }) => {
        const item = document.createElement("div")
        item.className = "sidebarItem sidebarItemCarpeta"
        item.textContent = materia
        item.dataset.normativa = normativa
        item.dataset.materia = materia
        item.dataset.carpetaId = carpeta.id
        item.style.borderLeftColor = obtenerColorMateria(materia)

        if (normativaSeleccionada === normativa && materiaSeleccionada === materia) {
          item.classList.add("activa")
        }

        item.addEventListener("click", () => {
          normativaSeleccionada = normativa
          materiaSeleccionada = materia
          sidebar.querySelectorAll(".sidebarItem").forEach(i => i.classList.remove("activa"))
          item.classList.add("activa")
          mostrarArticulosDeMateria(normativa, materia, agrupado[normativa][materia])
        })

        activarArrastreMateria(item, lista, normativa)
        lista.appendChild(item)
      })
    }

    card.appendChild(lista)
    contenedor.appendChild(card)
  })
}

function mostrarArticulosDeMateria(normativa, materia, items) {
  sincronizarEdiciones()

  const contenedor = document.getElementById("contenidoArticulos")
  contenedor.innerHTML = ""

  const wrap = document.createElement("div")
  wrap.className = "materia-wrap"

  const tituloM = document.createElement("div")
  tituloM.className = "materia-title"
  tituloM.contentEditable = "true"
  tituloM.textContent = materia
  const colorMateria = obtenerColorMateria(materia)
  tituloM.style.color = colorMateria
  tituloM.style.setProperty("--materia-underline", colorMateria)

  // Nombre actual que se va actualizando en vivo
  let nombreActual = materia

  // Actualización en tiempo real sin perder foco
  tituloM.addEventListener("input", () => {
    const nuevo = tituloM.textContent.trim()
    if (nuevo === "" || nuevo === nombreActual) return

    articulos.forEach(a => {
      if (a.materia === nombreActual && a.normativa === normativa) {
        a.materia = nuevo
      }
    })

    renombrarMateriaEnCarpetas(normativa, nombreActual, nuevo)

    document.querySelectorAll(".sidebarItem").forEach(item => {
      if (item.dataset.materia === nombreActual && item.dataset.normativa === normativa) {
        item.dataset.materia = nuevo
        item.textContent = nuevo
      }
    })

    if (
      materiasOrden[normativa] &&
      typeof materiasOrden[normativa][nombreActual] === "number"
    ) {
      materiasOrden[normativa][nuevo] = materiasOrden[normativa][nombreActual]
      delete materiasOrden[normativa][nombreActual]
      guardarOrdenMaterias()
    }

    nombreActual = nuevo
    materiaSeleccionada = nuevo
    guardarLocal()
  })

  // Al salir del título se reordena todo
  tituloM.addEventListener("blur", () => {
    ordenarYMostrar()
  })

  const controles = document.createElement("div")
  controles.className = "materia-controls"

  const colorInput = document.createElement("input")
  colorInput.type = "color"
  colorInput.className = "materia-color-input"
  colorInput.value = colorMateria

  colorInput.addEventListener("click", e => e.stopPropagation())
  colorInput.addEventListener("pointerdown", e => e.stopImmediatePropagation())

  colorInput.addEventListener("input", () => {
    const nuevoColor = colorInput.value
    tituloM.style.color = nuevoColor
    tituloM.style.setProperty("--materia-underline", nuevoColor)

    const tarjetas = contenedor.querySelectorAll(".articulo")
    tarjetas.forEach(t => {
      if (t.dataset.materia === nombreActual) {
        t.style.borderLeftColor = nuevoColor
        const tituloArt = t.querySelector(".articulo-titulo")
        if (tituloArt) tituloArt.style.color = nuevoColor
      }
    })

    const sidebarItems = document.querySelectorAll(".sidebarItem")
    sidebarItems.forEach(si => {
      if (si.dataset.materia === nombreActual) {
        si.style.borderLeftColor = nuevoColor
      }
    })
  })

  colorInput.addEventListener("change", () => {
    const nuevoColor = colorInput.value
    articulos.forEach(a => {
      if (a.materia === nombreActual && a.normativa === normativa) a.color = nuevoColor
    })
    guardarLocal()
  })

  const borrarMat = document.createElement("button")
  borrarMat.className = "btn-small"
  borrarMat.textContent = "Borrar materia"
  borrarMat.onclick = () => borrarMateria(nombreActual, normativa)

  controles.appendChild(colorInput)
  controles.appendChild(borrarMat)

  wrap.appendChild(tituloM)
  wrap.appendChild(controles)
  contenedor.appendChild(wrap)

  const listaOrdenada = [...items].sort((a, b) => valorOrdenArticulo(a) - valorOrdenArticulo(b))

  listaOrdenada.forEach(a => {
    const box = document.createElement("div")
    box.className = "articulo"
    box.style.borderLeftColor = a.color
    box.dataset.id = a.id
    box.dataset.materia = a.materia

    const titulo = document.createElement("div")
    titulo.className = "articulo-titulo"
    titulo.style.color = a.color

    const tituloEditable = document.createElement("span")
    tituloEditable.className = "articulo-titulo-texto"
    tituloEditable.contentEditable = "true"
    tituloEditable.textContent =
      a.tituloPersonalizado && a.tituloPersonalizado.trim()
        ? a.tituloPersonalizado
        : `Artículo ${a.numero}`

    tituloEditable.addEventListener("click", e => e.stopPropagation())

    tituloEditable.addEventListener("input", () => {
      const nuevo = tituloEditable.textContent.trim()
      a.tituloPersonalizado = nuevo
      guardarLocal()
    })

    tituloEditable.addEventListener("blur", () => {
      if (!tituloEditable.textContent.trim()) {
        a.tituloPersonalizado = ""
        tituloEditable.textContent = `Artículo ${a.numero}`
        guardarLocal()
      }
    })

    const indicador = document.createElement("span")
    indicador.className = "indicador"
    indicador.textContent = "▶"

    titulo.appendChild(tituloEditable)
    titulo.appendChild(indicador)

    const norm = document.createElement("div")
    norm.style.fontSize = "13px"
    norm.style.color = "#6b7280"
    norm.textContent = a.normativa === "civil" ? "Normativa Código Civil" : "Normativa Código Penal"

    const contenido = document.createElement("div")
    contenido.className = "articulo-contenido"
    contenido.contentEditable = "true"
    const contenidoGuardado =
      a.contenidoHTML ??
      (a.contenidoEditado !== undefined && a.contenidoEditado !== null
        ? escaparComoHTML(a.contenidoEditado)
        : escaparComoHTML(a.contenido))

    contenido.innerHTML = contenidoGuardado

    contenido.addEventListener("input", () => {
      a.contenidoEditado = contenido.textContent
      a.contenidoHTML = contenido.innerHTML
      guardarLocal()
    })

    contenido.addEventListener("click", e => e.stopPropagation())

    const botonNota = document.createElement("button")
    botonNota.className = "btn-note"
    botonNota.textContent = "Agregar nota"
    botonNota.style.display = "none"

    const botonEliminarNota = document.createElement("button")
    botonEliminarNota.className = "btn-note-delete"
    botonEliminarNota.textContent = "Eliminar nota"
    botonEliminarNota.style.display = "none"

    const notaBox = document.createElement("textarea")
    notaBox.className = "nota-box"
    notaBox.value = a.nota || ""

    notaBox.addEventListener("click", e => e.stopPropagation())
    notaBox.addEventListener("input", () => {
      a.nota = notaBox.value
      guardarLocal()
      if (notaBox.value.trim() !== "") {
        botonEliminarNota.style.display = "inline-block"
      }
    })

    const resizer = document.createElement("div")
    resizer.className = "nota-resizer"

    resizer.addEventListener("mousedown", e => {
      e.stopPropagation()
      const startY = e.clientY
      const startHeight = notaBox.offsetHeight

      function mover(ev) {
        const nueva = startHeight + (ev.clientY - startY)
        notaBox.style.height = Math.max(nueva, 70) + "px"
      }

      function stop() {
        window.removeEventListener("mousemove", mover)
        window.removeEventListener("mouseup", stop)
      }

      window.addEventListener("mousemove", mover)
      window.addEventListener("mouseup", stop)
    })

    botonNota.addEventListener("click", e => {
      e.stopPropagation()
      notaBox.style.display = "block"
      resizer.style.display = "block"
      notaBox.focus()
    })

    botonEliminarNota.addEventListener("click", e => {
      e.stopPropagation()
      notaBox.value = ""
      a.nota = ""
      guardarLocal()
      notaBox.style.display = "none"
      resizer.style.display = "none"
      botonEliminarNota.style.display = "none"
    })

    if (a.nota && a.nota.trim() !== "") {
      notaBox.style.display = "block"
      resizer.style.display = "block"
      botonEliminarNota.style.display = "inline-block"
    }

    const botonBorrar = document.createElement("button")
    botonBorrar.className = "btn-small"
    botonBorrar.textContent = "Eliminar artículo"
    botonBorrar.style.marginTop = "16px"
    botonBorrar.addEventListener("click", e => {
      e.stopPropagation()
      borrarArticulo(a.id)
    })

    box.appendChild(titulo)
    box.appendChild(norm)
    box.appendChild(contenido)
    box.appendChild(botonNota)
    box.appendChild(botonEliminarNota)
    box.appendChild(notaBox)
    box.appendChild(resizer)
    box.appendChild(botonBorrar)

    box.addEventListener("click", () => {
      const abierta = box.classList.toggle("abierto")

      if (abierta) {
        botonNota.style.display = "inline-block"
        if (a.nota && a.nota.trim() !== "") {
          notaBox.style.display = "block"
          resizer.style.display = "block"
          botonEliminarNota.style.display = "inline-block"
        }
      } else {
        botonNota.style.display = "none"
        if (!a.nota || a.nota.trim() === "") {
          notaBox.style.display = "none"
          resizer.style.display = "none"
          botonEliminarNota.style.display = "none"
        }
      }
    })

    activarArrastreArticulo(box, normativa, materia)

    contenedor.appendChild(box)
  })
}

function guardarLocal() {
  localStorage.setItem("articulosGuardados", JSON.stringify(articulos))
}

function sincronizarEdiciones() {
  document.querySelectorAll(".articulo").forEach(box => {
    const id = Number(box.dataset.id)
    const articulo = articulos.find(a => a.id === id)
    if (!articulo) return

    const tituloEl = box.querySelector(".articulo-titulo-texto")
    if (tituloEl) articulo.tituloPersonalizado = tituloEl.textContent.trim()

    const contenidoEl = box.querySelector(".articulo-contenido")
    if (contenidoEl) {
      articulo.contenidoEditado = contenidoEl.textContent
      articulo.contenidoHTML = contenidoEl.innerHTML
    }

    const notaEl = box.querySelector(".nota-box")
    if (notaEl) articulo.nota = notaEl.value
  })

  guardarLocal()
}

function borrarArticulo(id) {
  articulos = articulos.filter(a => a.id !== id)
  guardarLocal()
  ordenarYMostrar()
}

function borrarMateria(nombre, normativa = null) {
  articulos = articulos.filter(a => {
    if (normativa) return a.materia !== nombre || a.normativa !== normativa
    return a.materia !== nombre
  })
  if (normativa) {
    removerMateriaDeCarpetas(normativa, nombre)
  } else {
    ["civil", "penal"].forEach(norm => removerMateriaDeCarpetas(norm, nombre))
  }
  if (normativa && materiasOrden[normativa]) {
    delete materiasOrden[normativa][nombre]
    guardarOrdenMaterias()
  }
  guardarLocal()
  normativaSeleccionada = null
  materiaSeleccionada = null
  ordenarYMostrar()
}

function abrirModalCarpeta(id = null) {
  if (!modalCarpeta || !modalCarpetaTitulo || !inputNombreCarpeta) return
  carpetaEnEdicionId = id
  const carpeta = carpetas.find(c => c.id === id)
  modalCarpetaTitulo.textContent = id ? "Editar carpeta" : "Nueva carpeta"
  inputNombreCarpeta.value = carpeta?.nombre || ""
  modalCarpeta.classList.add("visible")
  setTimeout(() => inputNombreCarpeta.focus(), 50)
}

function cerrarModalCarpeta() {
  if (!modalCarpeta || !inputNombreCarpeta) return
  modalCarpeta.classList.remove("visible")
  inputNombreCarpeta.value = ""
  carpetaEnEdicionId = null
}

function confirmarCarpeta() {
  if (!inputNombreCarpeta) return
  const nombre = inputNombreCarpeta.value.trim()
  if (!nombre) {
    inputNombreCarpeta.focus()
    return
  }

  if (carpetaEnEdicionId) {
    carpetas = carpetas.map(c => (c.id === carpetaEnEdicionId ? { ...c, nombre } : c))
  } else {
    carpetas = [
      ...carpetas,
      { id: `carpeta-${Date.now()}-${Math.random().toString(16).slice(2)}`, nombre, materias: [] }
    ]
  }

  guardarCarpetas()
  cerrarModalCarpeta()
  ordenarYMostrar()
}

function eliminarCarpeta(id) {
  carpetas = carpetas.filter(c => c.id !== id)
  guardarCarpetas()
  ordenarYMostrar()
}

inicializarMenuContextual()
renderDocumentos()
ordenarYMostrar()

window.addEventListener("beforeunload", sincronizarEdiciones)
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden") sincronizarEdiciones()
})