import os
import re
import tempfile
from datetime import datetime
import openai
from openai import OpenAI
import speech_recognition as sr
import pygame
from dotenv import load_dotenv
from elevenlabs.client import ElevenLabs
import openpyxl

load_dotenv()  # Carga ELEVENLABS_API_KEY (y demás claves) desde API/.env

# --- Configuración general -------------------------------------------
# URL del servidor local de LM Studio (pestaña "Local Server", botón
# "Start Server"). El puerto por defecto es 1234.
LMSTUDIO_BASE_URL = "http://localhost:1234/v1"

MAX_PREGUNTAS = 10
MAX_TOKENS_RESPUESTA = 600
IDIOMA_VOZ = "es-ES"  # Cambia a "es-MX", "es-AR", "es-CO"... si prefieres otro acento

# --- Configuración de la voz (ElevenLabs) -----------------------------
ELEVENLABS_API_KEY = os.getenv("ELEVENLABS_API_KEY")
# ID de la voz a usar. Puedes buscar más voces (y su ID) en
# https://elevenlabs.io/app/voice-library. Por defecto se usa "Rachel",
# una voz predefinida de ElevenLabs; cámbiala poniendo ELEVENLABS_VOICE_ID
# en el .env si prefieres otra.
ELEVENLABS_VOICE_ID = os.getenv("ELEVENLABS_VOICE_ID", "21m00Tcm4TlvDq8ikWAM")
# Modelo multilingüe: necesario para que la voz suene bien en español.
ELEVENLABS_MODEL_ID = "eleven_multilingual_v2"
ELEVENLABS_OUTPUT_FORMAT = "mp3_44100_128"

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


def conectar_lmstudio():
    """Crea el cliente apuntando al servidor local de LM Studio."""
    return OpenAI(base_url=LMSTUDIO_BASE_URL, api_key="lm-studio")


def seleccionar_modelo(cliente):
    """Detecta los modelos cargados en LM Studio y deja elegir uno si hay varios."""
    try:
        modelos = cliente.models.list()
    except openai.APIConnectionError as e:
        raise RuntimeError(
            "No se pudo conectar con LM Studio. Abre LM Studio, ve a la "
            "pestaña 'Local Server', carga un modelo y pulsa 'Start "
            f"Server' (debe quedar escuchando en {LMSTUDIO_BASE_URL})."
        ) from e

    ids_modelos = [modelo.id for modelo in modelos.data]
    if not ids_modelos:
        raise RuntimeError(
            "LM Studio no tiene ningún modelo cargado. Carga uno desde "
            "la pestaña 'Local Server' antes de iniciar la consulta."
        )

    if len(ids_modelos) == 1:
        print(f"\n✅ Usando el modelo cargado en LM Studio: {ids_modelos[0]}")
        return ids_modelos[0]

    print("=" * 60)
    print("  Hay varios modelos cargados en LM Studio. Elige uno:")
    print("=" * 60)
    for indice, modelo_id in enumerate(ids_modelos, start=1):
        print(f"  [{indice}] {modelo_id}")

    while True:
        eleccion = input(f"\nEscribe el número del modelo (1-{len(ids_modelos)}): ").strip()
        if eleccion.isdigit() and 1 <= int(eleccion) <= len(ids_modelos):
            return ids_modelos[int(eleccion) - 1]
        print("⚠️ Opción no válida, intenta de nuevo.")


def en_jupyter():
    """Detecta si el código se está ejecutando dentro de un notebook Jupyter."""
    try:
        from IPython import get_ipython
        shell = get_ipython()
        return shell is not None and "ZMQInteractiveShell" in str(type(shell))
    except ImportError:
        return False


_cliente_voz = None


def configurar_tts():
    """Crea (una sola vez) el cliente de ElevenLabs y el mezclador de audio."""
    global _cliente_voz

    if _cliente_voz is None:
        if not ELEVENLABS_API_KEY:
            raise RuntimeError(
                "Falta la variable de entorno ELEVENLABS_API_KEY. Añádela "
                "al archivo API/.env con tu clave de ElevenLabs (la "
                "encuentras en https://elevenlabs.io/app/settings/api-keys)."
            )
        pygame.mixer.init()
        _cliente_voz = ElevenLabs(api_key=ELEVENLABS_API_KEY)

    return _cliente_voz


def hablar(texto):
    """Hace que el Dr. Joy diga el texto en voz alta (con ElevenLabs) y lo muestra en pantalla."""
    print(f"\nDr. Joy: {texto}")
    cliente_voz = configurar_tts()

    audio = cliente_voz.text_to_speech.convert(
        voice_id=ELEVENLABS_VOICE_ID,
        model_id=ELEVENLABS_MODEL_ID,
        text=texto,
        output_format=ELEVENLABS_OUTPUT_FORMAT,
    )
    audio_bytes = b"".join(audio)

    ruta_audio = os.path.join(tempfile.gettempdir(), "dr_joy_voz.mp3")
    with open(ruta_audio, "wb") as archivo_audio:
        archivo_audio.write(audio_bytes)

    if en_jupyter():
        from IPython.display import Audio, display

        display(Audio(ruta_audio, autoplay=True))
    else:
        pygame.mixer.music.load(ruta_audio)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            pygame.time.wait(100)


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


def quitar_pensamiento(texto):
    """Quita el razonamiento interno que algunos modelos locales devuelven
    entre etiquetas <think>...</think> (Qwen3, DeepSeek-R1...) antes de su
    respuesta final, para que no interfiera al buscar RESUMEN/GRAVEDAD.
    """
    sin_pensamiento = re.sub(r"<think>.*?</think>", "", texto, flags=re.DOTALL | re.IGNORECASE)
    # Si se cortó a mitad del razonamiento (sin etiqueta de cierre, por
    # haberse quedado sin tokens), descarta también ese resto sin cerrar.
    sin_pensamiento = re.split(r"<think>", sin_pensamiento, flags=re.IGNORECASE)[0]
    return sin_pensamiento.strip()


def generar_resumen_y_gravedad(cliente, modelo_id, historial):
    """Pide al modelo un resumen de la dolencia y una nota de gravedad (1-5).

    1 significa que no es una urgencia y 5 que el paciente debe ser
    atendido de manera inmediata. Devuelve una tupla
    (resumen, gravedad, tokens_entrada, tokens_salida).
    """
    if not historial:
        return "Sin síntomas registrados.", 1, 0, 0

    mensajes = (
        [{"role": "system", "content": SYSTEM_PROMPT}]
        + historial
        + [
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
    )

    try:
        respuesta = cliente.chat.completions.create(
            model=modelo_id,
            # Los modelos "razonadores" locales (Qwen3, DeepSeek-R1...)
            # gastan tokens pensando antes de responder; con un límite
            # bajo se quedan sin tokens y el contenido sale vacío.
            max_tokens=500,
            messages=mensajes,
        )
    except openai.OpenAIError:
        return "No se pudo generar el resumen automático.", 1, 0, 0

    texto = quitar_pensamiento(respuesta.choices[0].message.content or "")

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
    tokens_entrada = uso.prompt_tokens if uso else 0
    tokens_salida = uso.completion_tokens if uso else 0

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
    cliente = conectar_lmstudio()
    modelo_id = seleccionar_modelo(cliente)

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
            respuesta = cliente.chat.completions.create(
                model=modelo_id,
                max_tokens=MAX_TOKENS_RESPUESTA,
                messages=[{"role": "system", "content": SYSTEM_PROMPT}] + historial,
            )
        except openai.OpenAIError as e:
            print(f"\n⚠️ Hubo un problema al contactar con LM Studio: {e}")
            break

        texto_respuesta = respuesta.choices[0].message.content or ""

        hablar(texto_respuesta)

        # --- Consumo de tokens de este turno (inferencia local, sin coste) ---
        uso = respuesta.usage
        tokens_entrada = uso.prompt_tokens if uso else 0
        tokens_salida = uso.completion_tokens if uso else 0
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
        generar_resumen_y_gravedad(cliente, modelo_id, historial)
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
    print(f"  Modelo usado                : {modelo_id}")
    print(f"  Tokens de entrada (prompt)  : {total_tokens_entrada}")
    print(f"  Tokens de salida (respuesta): {total_tokens_salida}")
    print(f"  Tokens totales               : {total_tokens_entrada + total_tokens_salida}")
    print("=" * 60)


if __name__ == "__main__":
    chat_con_medico()