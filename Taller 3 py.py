import re
import tkinter as tk
from collections import Counter
from docx import Document
import matplotlib 
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from wordcloud import WordCloud


STOPWORDS = {
    
    "el","la","los","las","un","una","unos","unas","al","del",
    
   
    "este","esta","estos","estas","ese","esa","esos","esas",
    "aquel","aquella","aquellos","aquellas",
    "mi","mis","tu","tus","su","sus",
    "nuestro","nuestra","nuestros","nuestras",
    
   
    "mucho","mucha","muchos","muchas","poco","poca","pocos","pocas",
    "bastante","demasiado","varios","varias",
    
   
    "bien","mal","mejor","peor","siempre","nunca","ya",
    "aqui","aquí","alli","allí","alla","allá",
    "hoy","ayer","mañana","mas","más","menos",
    "tambien","también","solo","sólo","se","lo"
    
    
    "y","e","o","u","ni","que","porque","pues","como","cuando",
    "aunque","pero","sino","si","mientras","donde",
    "a","ante","bajo","con","contra","de","desde",
    "durante","en","entre","hacia","hasta","para",
    "por","según","sin","sobre","tras"
}

def contar_palabras_docx(nombre_archivo, top_n=10):
    try:
        doc = Document(nombre_archivo)
        texto = ""

       
        for parrafo in doc.paragraphs:
            texto += parrafo.text + " "

     
        texto = texto.lower()

       
        palabras = re.findall(r'\b[\wáéíóúüñ]+\b', texto, flags=re.UNICODE)

       
        palabras_filtradas = [
            p for p in palabras
            if p not in STOPWORDS and not p.endswith("mente")
        ]

        
        conteo = Counter(palabras_filtradas)

        
        for palabra, frecuencia in conteo.most_common():
            print(f"{palabra}: {frecuencia}")

    except FileNotFoundError:   
        print(f"El archivo '{nombre_archivo}' no fue encontrado.")
    except Exception as e:
        print(f"Ocurrió un error: {e}")   
        top_palabras = conteo.most_common(top_n)
        if top_palabras:
            palabras_grafico, frecuencias = zip(*top_palabras)

        plt.figure(figsize=(10, 6))
        plt.bar(palabras_grafico, frecuencias, color="skyblue", edgecolor="black")
        plt.title(f"Top {top_n} palabras más frecuentes", fontsize=14)
        plt.xlabel("Palabras", fontsize=12)
        plt.ylabel("Frecuencia", fontsize=12)
        plt.xticks(rotation=45)
        plt.tight_layout()

        plt.savefig(r"C:\Users\b12\Downloads\grafico_palabras.png", dpi=300, bbox_inches="tight")
         
        if palabras:
            wc = WordCloud(
                width=800,
                height=400,
                background_color="white",
                colormap="viridis",
                stopwords=STOPWORDS
            ).generate(" ".join(palabras))

            plt.figure(figsize=(10, 6))
            plt.imshow(wc, interpolation="bilinear")
            plt.axis("off")
            plt.title("Nube de Palabras", fontsize=16)
            plt.savefig(r"C:\Users\b12\Downloads\nube_palabras.png", dpi=300, bbox_inches="tight")
            plt.close()

    except FileNotFoundError:
        print(f"El archivo '{nombre_archivo}' no fue encontrado.")
    except Exception as e:
        print(f"Ocurrió un error: {e}")




contar_palabras_docx(r"C:\Users\b12\Downloads\documento.docx")




