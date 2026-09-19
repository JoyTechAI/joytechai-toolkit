import os
import re
import shutil
from datetime import datetime
import anthropic
import speech_recognition as sr
import pyttsx3
from dotenv import load_dotenv
import openpyxl

load_dotenv()  # Carga ANTHROPIC_API_KEY (y demás claves) desde API/.env


def _asegurar_flac_en_path():
    """Pone a disposición el conversor FLAC que necesita `recognize_google`.

    Desde la versión 3.10, la librería `speech_recognition` dejó de
    localizar automáticamente el binario de FLAC que trae empaquetado (por
    seguridad) y ahora exige tenerlo en el PATH del sistema. Además, en
    Windows busca un archivo llamado exactamente "flac" (sin extensión
    ".exe"), así que no basta con tener flac.exe en el PATH. En vez de
    obligarte a instalar FLAC aparte, copiamos (una sola vez) el binario
    que ya incluye el propio paquete a una carpeta local, con el nombre
    exacto que espera, y añadimos esa carpeta al PATH de este proceso.
    """
    if os.name != "nt":
        return  # en Linux/Mac se espera tener `flac` instalado por el sistema

    flac_origen = os.path.join(os.path.dirname(sr.__file__), "flac-win32.exe")
    if not os.path.exists(flac_origen):
        return

    carpeta_bin = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".bin")
    flac_destino = os.path.join(carpeta_bin, "flac")  # sin extensión: así lo busca la librería

    if not os.path.exists(flac_destino):
        os.makedirs(carpeta_bin, exist_ok=True)
        shutil.copyfile(flac_origen, flac_destino)

    if carpeta_bin not in os.environ.get("PATH", "").split(os.pathsep):
        os.environ["PATH"] = carpeta_bin + os.pathsep + os.environ.get("PATH", "")


_asegurar_flac_en_path()

# --- Configuración general -------------------------------------------
CLAUDE_MODEL = "claude-sonnet-5"

MAX_PREGUNTAS = 10
MAX_TOKENS_RESPUESTA = 600
IDIOMA_VOZ = "es-ES"  # Cambia a "es-MX", "es-AR", "es-CO"... si prefieres otro acento

# --- Configuración de la voz (pyttsx3: offline, gratis, sin API key) --
# Velocidad de habla en palabras por minuto.
TTS_VELOCIDAD = int(os.getenv("TTS_VELOCIDAD", "175"))
# ID de una voz concreta instalada en Windows (opcional, ver
# listar_voces_disponibles() más abajo para encontrar el ID). Si se deja
# vacío, se intenta detectar automáticamente una voz en español entre
# las instaladas en el sistema; si no hay ninguna, se usa la voz por
# defecto de Windows.
TTS_VOZ_ID = os.getenv("TTS_VOZ_ID", "")

PALABRAS_SALIDA = ("salir", "adiós", "adios", "terminar", "hasta luego", "para")
PALABRAS_AFIRMATIVAS = ("sí", "si", "claro", "vale", "afirmativo", "quiero", "deseo", "ok")

# Si el paciente dice más de esta cantidad de veces que el dato
# reconocido por voz es incorrecto, se le deja escribirlo por teclado.
INTENTOS_VOZ_ANTES_DE_ESCRIBIR = 2

NUMEROS_EN_PALABRAS = {
    "cero": "0", "uno": "1", "una": "1", "dos": "2", "tres": "3",
    "cuatro": "4", "cinco": "5", "seis": "6", "siete": "7",
    "ocho": "8", "nueve": "9", "diez": "10",
}

RUTA_EXCEL_PACIENTES = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "pacientes.xlsx"
)

SYSTEM_PROMPT = """\
Eres el Dr. Joy, un médico de cabecera muy amable, cercano y empático.
Hablas en español, con un tono cálido, paciente y tranquilizador, como
un médico de confianza que se toma su tiempo para escuchar.

Esta conversación ocurre por VOZ, así que responde de forma breve y
natural, como en una charla hablada (evita listas largas, markdown o
párrafos extensos; usa frases cortas y claras, como hablaría un médico
en consulta).

Tu forma de trabajar:
- Escuchas con atención y haces preguntas claras para entender mejor
  los síntomas o dudas de la persona (una pregunta de seguimiento a
  la vez, no la satures).
- Explicas las cosas en lenguaje sencillo, evitando tecnicismos
  innecesarios.
- Ofreces orientación general de salud y bienestar.

Límites importantes que SIEMPRE respetas:
- No emites diagnósticos definitivos ni recetas con dosis de
  medicamentos. Puedes hablar de forma general sobre posibles causas
  o cuidados básicos, pero siempre recomendando confirmarlo con un
  profesional sanitario presencial.
- Si detectas señales de urgencia (dolor intenso, dificultad para
  respirar, pensamientos de hacerse daño, etc.), recomiendas con
  claridad y calma buscar atención médica inmediata o los servicios
  de emergencia.
- Dejas claro, de forma natural y sin sonar robótico, que eres un
  asistente virtual y no sustituyes una consulta médica real.
"""


def conectar_claude():
    """Crea el cliente de la API de Claude (Anthropic)."""
    if not os.getenv("ANTHROPIC_API_KEY"):
        raise RuntimeError(
            "Falta la variable de entorno ANTHROPIC_API_KEY. Añádela al "
            "archivo API/.env con tu clave de la API de Claude (la "
            "encuentras en https://console.anthropic.com/settings/keys)."
        )
    return anthropic.Anthropic()


_motor_voz = None


def listar_voces_disponibles():
    """Imprime las voces instaladas en Windows y su ID, para elegir una en TTS_VOZ_ID."""
    motor = pyttsx3.init()
    for voz in motor.getProperty("voices"):
        print(f"  {voz.id}\n    nombre: {voz.name}")


def configurar_tts():
    """Crea (una sola vez) el motor de texto a voz (pyttsx3: offline y gratis)."""
    global _motor_voz

    if _motor_voz is None:
        _motor_voz = pyttsx3.init()
        _motor_voz.setProperty("rate", TTS_VELOCIDAD)

        voz_elegida = TTS_VOZ_ID
        if not voz_elegida:
            # Sin voz configurada explícitamente: busca una en español
            # entre las voces instaladas en el sistema.
            for voz in _motor_voz.getProperty("voices"):
                if "spanish" in voz.name.lower() or "español" in voz.name.lower():
                    voz_elegida = voz.id
                    break

        if voz_elegida:
            _motor_voz.setProperty("voice", voz_elegida)

    return _motor_voz


def hablar(texto):
    """Hace que el Dr. Joy diga el texto en voz alta (con pyttsx3, offline y gratis)."""
    print(f"\nDr. Joy: {texto}")
    motor = configurar_tts()
    motor.say(texto)
    motor.runAndWait()


def escuchar(reconocedor, microfono):
    """Escucha por el micrófono y devuelve el texto reconocido, o None si falla."""
    with microfono as fuente:
        reconocedor.adjust_for_ambient_noise(fuente, duration=0.5)
        print("\n🎙️  Te escucho... (habla ahora)")
        try:
            audio = reconocedor.listen(fuente, timeout=8, phrase_time_limit=25)
        except sr.WaitTimeoutError:
            print("⚠️ No detecté que hablaras. Inténtalo de nuevo.")
            return None

    try:
        texto = reconocedor.recognize_google(audio, language=IDIOMA_VOZ)
        print(f"Tú: {texto}")
        return texto
    except sr.UnknownValueError:
        print("⚠️ No entendí lo que dijiste, ¿puedes repetirlo?")
        return None
    except sr.RequestError as e:
        print(f"⚠️ Error al conectar con el servicio de reconocimiento de voz: {e}")
        return None


def pedir_dato_por_voz(pregunta, reconocedor, microfono, intentos=3):
    """Hace una pregunta por voz y devuelve la respuesta reconocida.

    Reintenta hasta `intentos` veces si no se entiende nada. Si se
    agotan los intentos, devuelve None.
    """
    hablar(pregunta)
    for intento in range(intentos):
        respuesta = escuchar(reconocedor, microfono)
        if respuesta:
            return respuesta
        if intento < intentos - 1:
            hablar("Perdona, no te he entendido bien. ¿Puedes repetirlo?")
    return None


def es_respuesta_afirmativa(texto):
    """Comprueba si una respuesta hablada es un 'sí' en sentido amplio."""
    texto = (texto or "").lower()
    return any(palabra in texto for palabra in PALABRAS_AFIRMATIVAS)


def extraer_numero(texto):
    """Convierte una respuesta hablada en una cadena de solo dígitos.

    Sustituye los números dichos con letras ("uno", "dos"...) por su
    dígito correspondiente y descarta el resto de palabras, para que
    la tarjeta sanitaria quede guardada como cifras (1, 2, 3...) y no
    como texto (uno, dos, tres...).
    """
    if not texto:
        return ""

    digitos = []
    for palabra in re.findall(r"[\wáéíóúñ]+", texto.lower()):
        if palabra.isdigit():
            digitos.append(palabra)
        elif palabra in NUMEROS_EN_PALABRAS:
            digitos.append(NUMEROS_EN_PALABRAS[palabra])

    return "".join(digitos)


def pedir_nombre_confirmado(reconocedor, microfono, intentos_voz=INTENTOS_VOZ_ANTES_DE_ESCRIBIR):
    """Pide el nombre completo del paciente y lo confirma con él antes de seguir.

    Lee en voz alta el nombre entendido y pregunta si está bien
    escrito; si el paciente dice que no, vuelve a pedirlo por voz.
    Si dice que es incorrecto más de `intentos_voz` veces, se le deja
    escribirlo por teclado.
    """
    nombre = pedir_dato_por_voz(
        "¿Cuál es tu nombre completo? Dime tu nombre y tu apellido, por favor.",
        reconocedor,
        microfono,
    ) or "no proporcionado"

    for intento in range(intentos_voz + 1):
        confirmacion = pedir_dato_por_voz(
            f"He entendido tu nombre como: {nombre}. ¿Es correcto?",
            reconocedor,
            microfono,
        )
        if es_respuesta_afirmativa(confirmacion):
            return nombre

        if intento < intentos_voz:
            nombre = pedir_dato_por_voz(
                "De acuerdo, dime tu nombre completo de nuevo, por favor.",
                reconocedor,
                microfono,
            ) or nombre

    hablar("Vamos a probar de otra forma. Escribe tu nombre completo, por favor.")
    nombre_escrito = input("✍️  Escribe tu nombre completo: ").strip()
    return nombre_escrito or nombre


def pedir_numero_tarjeta_confirmado(
    reconocedor, microfono, intentos_voz=INTENTOS_VOZ_ANTES_DE_ESCRIBIR
):
    """Pide el número de la tarjeta sanitaria por voz y lo confirma con el paciente.

    El número se guarda siempre como dígitos (1, 2, 3...), nunca como
    palabras (uno, dos, tres...). Si el paciente dice que el número
    entendido es incorrecto más de `intentos_voz` veces, se le deja
    escribirlo por teclado.
    """
    numero = extraer_numero(
        pedir_dato_por_voz(
            "Gracias. Ahora dime el número de tu tarjeta sanitaria, por "
            "favor, número por número.",
            reconocedor,
            microfono,
        )
    )

    for intento in range(intentos_voz + 1):
        if numero:
            confirmacion = pedir_dato_por_voz(
                f"He entendido el número de tu tarjeta como: {numero}. ¿Es correcto?",
                reconocedor,
                microfono,
            )
            if es_respuesta_afirmativa(confirmacion):
                return numero

        if intento < intentos_voz:
            numero = extraer_numero(
                pedir_dato_por_voz(
                    "De acuerdo, dime el número de tu tarjeta de nuevo, "
                    "número por número, por favor.",
                    reconocedor,
                    microfono,
                )
            )

    hablar(
        "Vamos a probar de otra forma. Escribe el número de tu tarjeta "
        "sanitaria, por favor."
    )
    numero_escrito = "".join(
        caracter for caracter in input("✍️  Escribe el número de tu tarjeta sanitaria: ")
        if caracter.isdigit()
    )
    return numero_escrito or "no proporcionado"


def generar_resumen_y_gravedad(cliente, historial):
    """Pide al modelo un resumen de la dolencia y una nota de gravedad (1-5).

    1 significa que no es una urgencia y 5 que el paciente debe ser
    atendido de manera inmediata. Devuelve una tupla
    (resumen, gravedad, tokens_entrada, tokens_salida).
    """
    if not historial:
        return "Sin síntomas registrados.", 1, 0, 0

    mensajes = historial + [
        {
            "role": "user",
            "content": (
                "Basándote en toda la conversación anterior, responde "
                "ÚNICAMENTE con este formato exacto y nada más:\n"
                "RESUMEN: <una frase breve y clínica con el motivo de "
                "consulta y los síntomas principales>\n"
                "GRAVEDAD: <un número del 1 al 5, donde 1 significa que "
                "no es una urgencia y 5 que el paciente debe recibir "
                "atención médica de forma inmediata>"
            ),
        }
    ]

    try:
        respuesta = cliente.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=500,
            system=SYSTEM_PROMPT,
            thinking={"type": "disabled"},
            output_config={"effort": "low"},
            messages=mensajes,
        )
    except anthropic.APIError:
        return "No se pudo generar el resumen automático.", 1, 0, 0

    texto = next((bloque.text for bloque in respuesta.content if bloque.type == "text"), "")

    resumen_match = re.search(r"RESUMEN:\s*(.+)", texto)
    gravedad_match = re.search(r"GRAVEDAD:\s*([1-5])", texto)

    if resumen_match:
        resumen = resumen_match.group(1).strip()
    elif texto:
        resumen = texto
    else:
        # El modelo no devolvió nada útil: se arma un resumen mínimo a
        # partir de lo que dijo el paciente, para no dejar el campo vacío.
        mensajes_paciente = [m["content"] for m in historial if m["role"] == "user"]
        resumen = "Síntomas mencionados por el paciente: " + " / ".join(mensajes_paciente)

    gravedad = int(gravedad_match.group(1)) if gravedad_match else 1

    uso = respuesta.usage
    tokens_entrada = uso.input_tokens if uso else 0
    tokens_salida = uso.output_tokens if uso else 0

    return resumen, gravedad, tokens_entrada, tokens_salida


def guardar_paciente_excel(
    nombre, numero_tarjeta, resumen_dolencia, gravedad, quiere_cita, ruta=RUTA_EXCEL_PACIENTES
):
    """Crea (si no existe) o actualiza pacientes.xlsx con una fila de esta consulta."""
    encabezados = [
        "Nombre",
        "Número de tarjeta",
        "Resumen de la dolencia",
        "Gravedad (1-5)",
        "Cita solicitada",
        "Fecha",
    ]

    if os.path.exists(ruta):
        libro = openpyxl.load_workbook(ruta)
        hoja = libro.active
    else:
        libro = openpyxl.Workbook()
        hoja = libro.active
        hoja.title = "Pacientes"
        hoja.append(encabezados)

    hoja.append(
        [
            nombre,
            numero_tarjeta,
            resumen_dolencia,
            gravedad,
            "Sí" if quiere_cita else "No",
            datetime.now().strftime("%Y-%m-%d %H:%M"),
        ]
    )
    libro.save(ruta)
    print(f"\n💾 Ficha del paciente guardada en: {ruta}")


def chat_con_medico():
    cliente = conectar_claude()

    reconocedor = sr.Recognizer()
    microfono = sr.Microphone()

    historial = []
    num_preguntas = 0
    total_tokens_entrada = 0
    total_tokens_salida = 0

    print("=" * 60)
    print("  Dr. Joy - Tu médico virtual amable (modo voz)")
    print(f"  (Puedes hacer hasta {MAX_PREGUNTAS} preguntas)")
    print("  Di 'salir' o 'adiós' para terminar antes de tiempo.")
    print("=" * 60)

    hablar("Hola, soy el Doctor Joy. Antes de empezar, necesito un par de datos.")

    nombre_paciente = pedir_nombre_confirmado(reconocedor, microfono)

    numero_tarjeta = pedir_numero_tarjeta_confirmado(reconocedor, microfono)

    print(f"\n📋 Paciente: {nombre_paciente}  |  Tarjeta sanitaria: {numero_tarjeta}")

    hablar(f"Perfecto, {nombre_paciente}. Cuéntame, ¿en qué puedo ayudarte hoy?")

    while num_preguntas < MAX_PREGUNTAS:
        print(f"\n[{num_preguntas + 1}/{MAX_PREGUNTAS}]")
        pregunta = escuchar(reconocedor, microfono)

        if pregunta is None:
            continue  # No se entendió nada: no consume una pregunta del límite

        if pregunta.lower().strip(" .,!¡¿?") in PALABRAS_SALIDA:
            hablar("Cuídate mucho. ¡Hasta pronto!")
            break

        historial.append({"role": "user", "content": pregunta})

        try:
            respuesta = cliente.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=MAX_TOKENS_RESPUESTA,
                system=SYSTEM_PROMPT,
                thinking={"type": "disabled"},
                output_config={"effort": "low"},
                messages=historial,
            )
        except anthropic.APIError as e:
            print(f"\n⚠️ Hubo un problema al contactar con la API de Claude: {e}")
            break

        texto_respuesta = next(
            (bloque.text for bloque in respuesta.content if bloque.type == "text"), ""
        )

        hablar(texto_respuesta)

        # --- Consumo de tokens de este turno ---
        uso = respuesta.usage
        tokens_entrada = uso.input_tokens if uso else 0
        tokens_salida = uso.output_tokens if uso else 0
        total_tokens_entrada += tokens_entrada
        total_tokens_salida += tokens_salida

        print(
            f"   🔹 Tokens de este turno → entrada: {tokens_entrada} | "
            f"salida: {tokens_salida} | total: {tokens_entrada + tokens_salida}"
        )
        print(
            f"   🔹 Acumulado de la conversación → entrada: {total_tokens_entrada} | "
            f"salida: {total_tokens_salida} | total: "
            f"{total_tokens_entrada + total_tokens_salida}"
        )

        historial.append({"role": "assistant", "content": texto_respuesta})
        num_preguntas += 1

    else:
        # Se alcanzó el límite de preguntas
        print("\n" + "-" * 60)
        print("Dr. Joy: Hemos llegado al límite de preguntas de esta consulta.")
        print("-" * 60)
        hablar(
            "Hemos llegado al límite de preguntas de esta consulta. "
            "Si los síntomas persisten o empeoran, te recomiendo acudir "
            "a un centro médico para una valoración en persona. Cuídate mucho.",
        )

    # --- Pregunta sobre pedir cita ---
    respuesta_cita = pedir_dato_por_voz(
        "Antes de terminar, ¿quieres pedir cita con el médico?",
        reconocedor,
        microfono,
    )
    quiere_cita = es_respuesta_afirmativa(respuesta_cita)

    if quiere_cita:
        hablar("Perfecto, ahora serás redirigido para que puedas reservar tu cita.")
    else:
        hablar("De acuerdo, no pediremos cita por ahora.")

    # --- Resumen de la dolencia, nota de gravedad y guardado en Excel ---
    resumen_dolencia, gravedad, tokens_resumen_entrada, tokens_resumen_salida = (
        generar_resumen_y_gravedad(cliente, historial)
    )
    total_tokens_entrada += tokens_resumen_entrada
    total_tokens_salida += tokens_resumen_salida

    # La nota de gravedad es uso interno: no se dice ni se muestra al
    # paciente, solo queda registrada en el Excel para el equipo médico.
    guardar_paciente_excel(
        nombre_paciente, numero_tarjeta, resumen_dolencia, gravedad, quiere_cita
    )

    # --- Resumen final de consumo de tokens (inferencia local, sin coste) ---
    print("\n" + "=" * 60)
    print("  RESUMEN DE CONSUMO")
    print(f"  Paciente                    : {nombre_paciente}")
    print(f"  Tarjeta sanitaria            : {numero_tarjeta}")
    print(f"  Resumen de la dolencia       : {resumen_dolencia}")
    print(f"  Cita solicitada              : {'Sí' if quiere_cita else 'No'}")
    print(f"  Modelo usado                : {CLAUDE_MODEL}")
    print(f"  Tokens de entrada (prompt)  : {total_tokens_entrada}")
    print(f"  Tokens de salida (respuesta): {total_tokens_salida}")
    print(f"  Tokens totales               : {total_tokens_entrada + total_tokens_salida}")
    print("=" * 60)


if __name__ == "__main__":
    chat_con_medico()