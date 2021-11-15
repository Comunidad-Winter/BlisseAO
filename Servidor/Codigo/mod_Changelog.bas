Attribute VB_Name = "Mod_ChangeLog"
'*******************************************
'Standelf                            |                      16/05/2010
'*******************************************
'Comando /Editame, Solo para GameMasters, _
    es medio lagoso.
'Enviamos el Nombre del item que está en el suelo.
'*******************************************
'Standelf                            |                      20/05/2010
'*******************************************
'Ahora los users pueden desparalizarse aunque no esten _
    paralizados, así gastan maná igual por Giles.
'Enviamos la información de daños y curaciones a Users _
    y NPCs sobre su cabeza.
'Dados "Más" fáciles.
'Quitado el tiempo de inicio al meditar
'Ahora se pueden controlar Variables generales desde el _
    mod_controlao, entre estas la cantidad de extracción _
    de trabajadores, experiencia, oro, etc.
'Enviamos Nombre del mapa al cambiar de mapa.
'Enviamos el Record de usuarios Online al enviar el /Online.
'*******************************************
'Standelf                            |                      22/05/2010
'*******************************************
'Ahora recibimos posición al tirar el item.
'Implementé Drag&Drop de objetos, ahora chekea la pos _
    donde se tira para que no sea inválida y no tira con _
    más de 4 tiles de distancia.
'Implemente el movimiento vía Drag&drop de items _
    dentro del Inventario.
'Agregada la función CaracterInvalido
'Chat Global
'Chat Faccionario
'Comando /Promedio
'Comando /Guardar
'Comando /Invasion, ahora los GMs podrán iniciar una _
    invasión de NPCs enviando /Invasion NumNPC.
'Fundar clan requiere 18 de carisma.
'*******************************************
'TonchitoZ                         |                      25/05/2010
'*******************************************
'Boveda de cuentas _
    10 items a depositar solo para usuarios VIP
'Creación de cuentas _
    Una cuenta común equivale a 5 PJ, 8 en el caso de los usuarios VIP.
'Logueo de cuentas _
    No se puede conectar 2 veces al mismo tiempo la cuenta*
'Borrado de PJs de cuentas _
    Se realizan copias de los PJs borrados de las cuentas y se ingresa la fecha y hora del mismo proceso. _
        Por si se desea recuperar. (No se si valdrá la pena hacer esto, pero por las dudas...)
'Experiencia y Oro desde un INI
'HappyHour para usuarios PREMIUM (Intervalo, Duración, Experiencia y Oro dados en Server.INI)
'*******************************************
'Standelf                            |                      26/05/2010
'*******************************************
'Ahora el sacerdote cura el veneno ¬¬
'*******************************************
'Standelf                            |                      27/05/2010
'*******************************************
'Sistema de Auras, Ahora a las Armaduras, Cascos _
    Escudos, o armas se les puede poner Aura=N desde _
    Su Dat.
'*******************************************
'Standelf                            |                      28/05/2010
'*******************************************
'Ahora mostramos cuando no se puede usar un item en _
    el inventario
'*******************************************
'TonchitoZ                         |                      30/05/2010
'*******************************************
'Ahora muestra el daño que le producen a los usuarios sobre su cabeza (Tanto NPC como Usuarios)
'*******************************************
'Standelf                            |                      31/05/2010
'*******************************************
'Actualizado el Drag&Drop, ahora al mover en inventario revisa los slots y suma los de la mochila.
'El drag&drop ahora informa si el slot donde deseamos mover es inválido
'*******************************************
'Standelf                            |                      01/05/2010
'*******************************************
'Terminado el Sistema de Partys en Screen ;)
'*******************************************
'Standelf                            |                      02/06/2010
'*******************************************
'Reparado el bug que quedó con las auras al morir ¬¬
'*******************************************
'Standelf                            |                      03/06/2010
'*******************************************
'Mejorado el Sistema de Auras
'Subastas!
'*******************************************
'Standelf                            |                      04/06/2010
'*******************************************
'Creado el NPC de Bovedas VIP, Al igual que el boveda normal pero con otro NPC Only Vips (H)
'*******************************************
'TonchitoZ                         |                      19/06/2010
'*******************************************
'Contador al mejor estilo Mortal Kombat 3,2,1, Figth! para eventos.
    '-No puedo lanzar hechizos mientras este en evento y este el contador ON
    '-No puedo ni moverme ni atacar mientras este en evento y este el contador ON
'*******************************************
'TonchitoZ                         |                      21/06/2010
'*******************************************
'Comienzo de sistema DathMatch (creación del mismo)
'*******************************************
'TonchitoZ                         |                      26/06/2010
'*******************************************
'Inscripción de usuarios
'*******************************************
'Standelf                         |                         22/07/2010
'*******************************************
'Reparado los GiveExp
'Reparado los multiplicadores de Experiencia y Oro
'Reparados los items de trabajadores Newbie
'Al crear un PJ se otorgan 100 de Oro
'Al crear un PJ Mago se otorgan 1 báculo en vez de 1 daga
'*******************************************

