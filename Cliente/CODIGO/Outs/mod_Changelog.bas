Attribute VB_Name = "mod_Changelog"
'*******************************************
'Standelf                            |                      12/05/2010
'*******************************************
'Borrado el #ConAlfaB (Futuro Engine DirectX8)
'Borrado el #ConMenuseConextuales (Inservible)
'Borrado el #SeguridadAlkon (Al RE pedo jaja, Seguridad BG Rlz!) Off: La seguridad de alkon ya se la recontra violaron hace rato igual xD.
'*******************************************
'Mannakia                            |                     13/05/2010
'*******************************************
'Engine B�sico DirectX8
'*******************************************
'Standelf                            |                      14/05/2010
'*******************************************
'DirectX8 Fixes
'Agregado variable de luz para tiles
'Agregado Variables para seleccionar color al renderizar
'Creada funci�n para Aplicar ColorValues a Arrays
'Creada funci�n para aplicar Longs a Arrays
'Light Engine, Creado originalmente por Mannakia, Adaptado al nuevo Engine.
'Elminado el winsocket 2 del frmMain, era de la seguridad Alkon ��
'Agregado Panel Tester, para poner ahi todo para testear y que solo nos joda un Label en el frmMain xD
'Fixeado Alpha de DirectX8, Iniciaci�n editada.
'Doble engine de luces, Ahora se podr� via argumentos opcionales Seleccionar Luces cuadradas o Redondas :O
'*******************************************
'Standelf                            |                      15/05/2010
'*******************************************
'Fonts por renderizado, Eliminado el D3DFonts.
'Engine de Estados b�sico, 5 estados.
'Modulo para controlar Opciones
'*******************************************
'Standelf                            |                      16/05/2010
'*******************************************
'Comando /EDITAME, solo v�lido para Game Masters.
'Recibimos el Nombre del Item que est� en el piso.
'Documentaci�n.
'*******************************************
'Standelf                            |                      20/05/2010
'*******************************************
'Recibimos y mostramos nombre del mapa en el frmMain.
'Dialogos suben desde la cabeza
'Textos Tienen efecto FadeIn-FadeOut
'*******************************************
'Standelf                            |                      21/05/2010
'*******************************************
'Sistema de seguridad de intervalos
'*******************************************
'Standelf                            |                      22/05/2010
'*******************************************
'Drag&Drop de items para tirarlos al suelo o acomodarlos dentro del inventario
'Fix Light Engine #1
'Chat Global
'Chat Faccionario
'Comando /Guardar
'Comando /Promedio
'Mostramos el tiempo de paralisis restante con una barrita y im�gen al final del screen
'Agregado el MiniMapa se mantiene siempre sobre el frmmain y puede moverse a gusto de el usuario tanto dentro como fuera de la pantalla principal.
'Comando /Invasion
'*******************************************
'TonchitoZ                         |                      25/05/2010
'*******************************************
'Un poco de limpieza en el Cliente, A�nf quedan resto de basuras.
'Modo de Vision de nombres (Siempre visibles, al recibir enfoque o nunca) ME COSTO DEMACIADO HACER ESTO!!!! XD
'Nose si poner que estube boludeando un toque con Flash para poder hacer una mejor interface de conectar
'Tambi�n me com� una pata y muslo asada con papas fritas
'Adem�s alguien me tiene que comprar m�s caf� �� porque ya me tome todo ��
'*******************************************
'Standelf                            |                      26/05/2010
'*******************************************
'Iniciado el PartyEngine, Luego lo terminar�.
'Con el /GM ahora se nos abre un formulario para explicar los motivos, denunciar y reportar bugs.
'Al crear PJ ya no se manda el nick en mayuscula
'Intervalos de seguridad asignados
'Creado M�dulo de seguridad General.
'Creado el DeInit del Engine Gr�fico
'Actualizada la Seguridad
'*******************************************
'Standelf                            |                      27/05/2010
'*******************************************
'Agregado Sistema de Auras
'Si el usuario es VIP ve una peque�a "animaci�n" debajo de �l
'En la lluvia hay relampagos (H)
'*******************************************
'Standelf                            |                      28/05/2010
'*******************************************
'Obtenemos el MAC del usuario
'Obtenemos el HD Serial del usuario
'Al mirar las caracteristicas de un clan si tiene web se puede acceder a esta.
'Ahora en la lluvia aparecen truenos (H)
'Mejoradas cosas de la seguridad
'Agregada la variable Client_Web que almacena la p�gina del juego
'Ahora mostramos cuando no se puede usar un item en el Inventario
'*******************************************
'TonchitoZ                         |                      29/05/2010
'*******************************************
'Nueva graficaci�n de la pantalla principal (frmMain)
'Integraci�n de interface del frmMain y adaptaci�n de los componentes
'Creaci�n de barra de Experiencia
'Ahora no parpadean los nombres de PJs del frmCuenta
'Limpieza del cliente
'Modificaci�n de clsFormMovementManager (Ahora tambien se puede mover desde un PictureBox, _
            sirve para el minimapa)
'Creaci�n de men� del juego (Faltan Manual y Foro que se pondr� cuando alla ^^)
'*******************************************
'Standelf                            |                      31/05/2010
'*******************************************
'Arregle los HORRORES de TonchitoZ (?)
'Ahora hablando en serio, arregle el bug que se genero al poner la barra de experiencia.
'Mejor� la configuraci�n del drag&drop
'El inventario ahora tiene su modo "din�mico"
'Puse el inventario gr�fico que tonchitoz no pudo =P
'Acomode posiciones del Main, Si soy muy perfeccionista y cada pixel lo acomodo :$
'*******************************************
'Standelf                            |                      01/06/2010
'*******************************************
'Actualizado el mod_client_settings
'Modificado todo el panel de Opciones'
'Terminado el PartyEngine!
'ACLARO ESTUBE TODA LA TARDE CON ESTA MIERDA �� PARTYS FEAS MAL ECHAS XD
'*******************************************
'Standelf                            |                      03/06/2010
'*******************************************
'Mejorado el Sistema de Auras
'Ahora al poner un Comando desconocido informa
'Subastas!
'*******************************************
'Standelf                            |                      04/06/2010
'*******************************************
'Edite el Inventario de Boveda VIP asi aparece tambi�n la cantidad que se tiene en el inventario.
'Consola de clan y general
'Tipos de Habla seleccionables
'Ahora s� guardamos la ultima cuenta ;)
'Documentaci�n
'*******************************************
'Standelf                            |                      05/06/2010
'*******************************************
'Reeditado el CharRender
'Al pasar el mouse por un PJ se pone Verde para remarcar
'Al pasar el mouse sobre un PJ y querer lanzar un hechizo nos aparecer� el "simbolo" si se tiene al lado.
'Los caspers ahora tienen colores por faccion y si son GMs
'El CharRender vuelve a ser uno solo
'Ahora al renderizar un GRH enviamos el color
'*******************************************
'TonchitoZ                         |                      05/06/2010
'*******************************************
'Agregada opci�n de tonalidad de MiniMapa
'Agregada opci�n para Activar o Desactivar el MiniMapa
'*******************************************
'Standelf                           |                      09/06/2010
'*******************************************
'Sistema de ScreenShots
'*******************************************
'Standelf                           |                      11/06/2010
'*******************************************
'Optimizado el Cliente, Limpieza de la iniciaci�n del Engine.
'Documentaci�n y Orden General.
'*******************************************
'Standelf                           |                      12/06/2010
'*******************************************
'Edite el Cargando y el SubMain ahora es mas lindo a la vista, mas informativo y tiene menos boludeces
'*******************************************
'Standelf                           |                      16/06/2010
'*******************************************
'Me acorde que este mod existia :S espero acordarme de todo lo que cambie, aver que onda:
'Particulas vbGore, Comunes y para perseguir al usuario ;)
'Proyectiles de vbGore, Faltan terminar se buguean por posici�n
'Creado el sistema de Log, sirve para dejar registrado todos los errores o sucesos importantes del sistema.
'*******************************************
'TonchitoZ                         |                      18/06/2010
'*******************************************
'Reparadas barras de Hambre y Sed que STANDELF dej� MAL pocisionadas (jum) vos me mandastes al frente antes =P
'*******************************************
'TonchitoZ                         |                      19/06/2010
'*******************************************
'Contador al mejor estilo Mortal Kombat 3,2,1, Figth! para eventos.
'*******************************************
'TonchitoZ                         |                      21/06/2010
'*******************************************
'Comienzo de DeathMatch
'*******************************************
'Standelf                         |                      11/07/2010
'*******************************************
'Cambie el sistema de subastas
'Borre modulos y cosas al pedo del cliente 0,22MB menos en el exe :)
'*******************************************
'Standelf             |           12/07/2010
'*******************************************
'Arreglado el problema del minimapa
'Agregado el logo "PREMIUM" en el crear personaje, panel cuenta y main del juego.
'*******************************************
