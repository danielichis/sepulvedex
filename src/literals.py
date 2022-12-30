import math
def getCardinal(N):
    Unidad=["", "primero", "segundo", "tercero",
            "cuarto", "quinto", "sexto", "septimo", "octavo", "noveno"]
    Decena=["", "decimo", "vigesimo", "trigesimo", "cuadragesimo",
            "quincuagesimo", "sexagesimo", "septuagesimo", "octogesimo", 
    "nonagesimo"]
    Centena=["", "centesimo", "ducentesimo", "tricentesimo",
            " cuadringentesimo", " quingentesimo", " sexcentesimo", 
    " septingentesimo"," octingentesimo", " noningentesimo"]
    u= N % 10
    d=int(math.floor(N/10))%10
    c=int(math.floor(N/100))
    if(N>=100): 
        return Centena[c]+" "+Decena[d]+" "+Unidad[u]
    else:
        if(N>=10):
            return Decena[d]+" "+Unidad[u]
        else:
            return Unidad[N].upper()


