import time


inicio = time.time()

duracao = round(time.time() - inicio,2)

print("Duração: "+ str(duracao))
x = 0
while x < 1000:
    x = x + 1
    print(x)

parada = duracao > 0.10
print("boleana>>>"+str(parada))
parada = True
while parada:
    duracao = round(time.time() - inicio,2)
    print("Pause iniciado aos :"+str(duracao))
    parada = round(time.time() - inicio,2) < 5.00
duracao = round(time.time() - inicio,2)
print("Pause Finalizado aos :"+str(duracao))  