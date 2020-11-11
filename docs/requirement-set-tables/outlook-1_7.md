| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[addHandlerAsync (EventType: Office. EventType \| String, Handler: any, Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Fügt einen Ereignishandler für ein unterstütztes Ereignis hinzu.|
||[organizer](/javascript/api/outlook/outlook.appointmentcompose#organizer)|Ruft den Organisator für die angegebene Besprechung ab.|
||[Wiederholung](/javascript/api/outlook/outlook.appointmentcompose#recurrence)|Ruft das Serienmuster eines Termins ab, oder legt dieses fest.|
||[removeHandlerAsync (EventType: Office. EventType \| String, Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#removehandlerasync-eventtype--options--callback--asyncresult-)|Entfernt die Ereignishandler für einen unterstützten Ereignistyp.|
||[seriesId](/javascript/api/outlook/outlook.appointmentcompose#seriesid)|Ruft die ID der Datenreihe ab, zu der eine Instanz gehört.|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[addHandlerAsync (EventType: Office. EventType \| String, Handler: any, Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Fügt einen Ereignishandler für ein unterstütztes Ereignis hinzu.|
||[Wiederholung](/javascript/api/outlook/outlook.appointmentread#recurrence)|Ruft das Serienmuster eines Termins ab.|
||[removeHandlerAsync (EventType: Office. EventType \| String, Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentread#removehandlerasync-eventtype--options--callback--asyncresult-)|Entfernt die Ereignishandler für einen unterstützten Ereignistyp.|
||[seriesId](/javascript/api/outlook/outlook.appointmentread#seriesid)|Ruft die ID der Datenreihe ab, zu der eine Instanz gehört.|
|[AppointmentTimeChangedEventArgs](/javascript/api/outlook/outlook.appointmenttimechangedeventargs)|[end](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#end)||
||[start](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#start)||
||[Typ](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#type)||
|[From](/javascript/api/outlook/outlook.from)|[getasync (Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <EmailAddressDetails> ) => void)](/javascript/api/outlook/outlook.from#getasync-options--callback--asyncresult-)|Ruft den from-Wert einer Nachricht ab.|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[addHandlerAsync (EventType: Office. EventType \| String, Handler: any, Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messagecompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Fügt einen Ereignishandler für ein unterstütztes Ereignis hinzu.|
||[Von](/javascript/api/outlook/outlook.messagecompose#from)|Ruft die E-Mail-Adresse des Absenders einer Nachricht ab.|
||[removeHandlerAsync (EventType: Office. EventType \| String, Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messagecompose#removehandlerasync-eventtype--options--callback--asyncresult-)|Entfernt die Ereignishandler für einen unterstützten Ereignistyp.|
||[seriesId](/javascript/api/outlook/outlook.messagecompose#seriesid)|Ruft die ID der Datenreihe ab, zu der eine Instanz gehört.|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[addHandlerAsync (EventType: Office. EventType \| String, Handler: any, Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messageread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Fügt einen Ereignishandler für ein unterstütztes Ereignis hinzu.|
||[Wiederholung](/javascript/api/outlook/outlook.messageread#recurrence)|Ruft das Serienmuster eines Termins ab.|
||[removeHandlerAsync (EventType: Office. EventType \| String, Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messageread#removehandlerasync-eventtype--options--callback--asyncresult-)|Entfernt die Ereignishandler für einen unterstützten Ereignistyp.|
||[seriesId](/javascript/api/outlook/outlook.messageread#seriesid)|Ruft die ID der Datenreihe ab, zu der eine Instanz gehört.|
|[Organisator](/javascript/api/outlook/outlook.organizer)|[getasync (Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <EmailAddressDetails> ) => void)](/javascript/api/outlook/outlook.organizer#getasync-options--callback--asyncresult-)|Ruft den Organizer-Wert eines Termins als {@Link Office. EmailAddressDetails | EmailAddressDetails}-Objekt|
|[RecipientsChangedEventArgs](/javascript/api/outlook/outlook.recipientschangedeventargs)|[changedRecipientFields](/javascript/api/outlook/outlook.recipientschangedeventargs#changedrecipientfields)||
||[Typ](/javascript/api/outlook/outlook.recipientschangedeventargs#type)||
|[RecipientsChangedFields](/javascript/api/outlook/outlook.recipientschangedfields)|[bcc](/javascript/api/outlook/outlook.recipientschangedfields#bcc)|Ruft ab, ob Empfänger im Feld **Bcc** geändert wurden.|
||[cc](/javascript/api/outlook/outlook.recipientschangedfields#cc)|Ruft ab, ob Empfänger im Feld **CC** geändert wurden.|
||[optionalAttendees](/javascript/api/outlook/outlook.recipientschangedfields#optionalattendees)|Ruft ab, ob optionale Teilnehmer geändert wurden.|
||[requiredAttendees](/javascript/api/outlook/outlook.recipientschangedfields#requiredattendees)|Ruft ab, ob erforderliche Teilnehmer geändert wurden.|
||[Ressourcen](/javascript/api/outlook/outlook.recipientschangedfields#resources)|Ruft ab, ob Ressourcen geändert wurden.|
||[An](/javascript/api/outlook/outlook.recipientschangedfields#to)|Ruft ab, ob Empfänger im Feld **an** geändert wurden.|
|[Serie](/javascript/api/outlook/outlook.recurrence)|[getasync (Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <Recurrence> ) => void)](/javascript/api/outlook/outlook.recurrence#getasync-options--callback--asyncresult-)|Gibt das aktuelle Serien Objekt einer Terminserie zurück.|
||[recurrenceProperties](/javascript/api/outlook/outlook.recurrence#recurrenceproperties)|Ruft die Eigenschaften der Terminserie ab oder legt Sie fest.|
||[recurrenceTimeZone](/javascript/api/outlook/outlook.recurrence#recurrencetimezone)|Ruft die Eigenschaften der Terminserie ab oder legt Sie fest.|
||[recurrenceType](/javascript/api/outlook/outlook.recurrence#recurrencetype)|Ruft den Typ der Terminserie ab oder legt ihn fest.|
||[serieszeit](/javascript/api/outlook/outlook.recurrence#seriestime)|{@Link Office. Series | Series Time}-Objekt ermöglicht es Ihnen, das Start-und Enddatum der Terminserie zu verwalten und|
||[setasync (recurrencePattern: Serie, Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.recurrence#setasync-recurrencepattern--options--callback--asyncresult-)|Legt das Serienmuster einer Terminserie fest.|
|[RecurrenceChangedEventArgs](/javascript/api/outlook/outlook.recurrencechangedeventargs)|[Wiederholung](/javascript/api/outlook/outlook.recurrencechangedeventargs#recurrence)||
||[Typ](/javascript/api/outlook/outlook.recurrencechangedeventargs#type)||
|[RecurrenceProperties](/javascript/api/outlook/outlook.recurrenceproperties)|[dayOfMonth](/javascript/api/outlook/outlook.recurrenceproperties#dayofmonth)|Stellt den Tag des Monats dar.|
||[dayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#dayofweek)|Stellt den Tag der Woche oder des Tages Typs dar, beispielsweise "Weekend Day" vs Wochentag.|
||[Tage](/javascript/api/outlook/outlook.recurrenceproperties#days)|Stellt die Gruppe von Tagen für diese Serie dar.|
||[firstDayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#firstdayofweek)|Stellt den ausgewählten ersten Tag der Woche dar, andernfalls ist der Standardwert in den Einstellungen des aktuellen Benutzers.|
||[Intervall](/javascript/api/outlook/outlook.recurrenceproperties#interval)|Stellt den Zeitraum zwischen Instanzen der gleichen Terminserie dar.|
||[Monat](/javascript/api/outlook/outlook.recurrenceproperties#month)|Stellt den Monat dar.|
||[weekNumber](/javascript/api/outlook/outlook.recurrenceproperties#weeknumber)|Stellt die Nummer der Woche im ausgewählten Monat dar, beispielsweise "First" für die erste Woche des Monats.|
|[RecurrenceTimeZone](/javascript/api/outlook/outlook.recurrencetimezone)|[name](/javascript/api/outlook/outlook.recurrencetimezone#name)|Stellt den Namen der Serien Zeitzone dar.|
||[Offset](/javascript/api/outlook/outlook.recurrencetimezone#offset)|Ganzzahliger Wert, der den Unterschied zwischen der lokalen Zeitzone und der UTC in Minuten angibt, an dem Datum, an dem die Besprechungs Reihe begann.|
|[SeriesTime](/javascript/api/outlook/outlook.seriestime)|[getduration ()](/javascript/api/outlook/outlook.seriestime#getduration--)|Ruft die Dauer in Minuten einer normalen Instanz in einer Terminserie ab.|
||[EndDate ()](/javascript/api/outlook/outlook.seriestime#getenddate--)|Ruft das Enddatum eines Serienmusters in der folgenden Datei ab.|
||[getEndTime()](/javascript/api/outlook/outlook.seriestime#getendtime--)|Ruft die Endzeit einer normalen Termin-oder Besprechungsanfrage Instanz eines Serienmusters in der Zeitzone ab, die der Benutzer oder|
||[getstartdate ()](/javascript/api/outlook/outlook.seriestime#getstartdate--)|Ruft das Startdatum eines Serienmusters in der folgenden|
||[StartTime ()](/javascript/api/outlook/outlook.seriestime#getstarttime--)|Ruft die Startzeit einer üblichen Termin Instanz eines Serienmusters in der Zeitzone ab, für die der Benutzer/das Add-in den festgelegten|
||[setduration (Minuten: Zahl)](/javascript/api/outlook/outlook.seriestime#setduration-minutes-)|Legt die Dauer aller Termine in einem Serienmuster fest.|
||[EndDate (Datum: Zeichenfolge)](/javascript/api/outlook/outlook.seriestime#setenddate-date-)|Legt den Endtermin einer Terminserie fest.|
||[EndDate (Jahr: Zahl, Monat: Zahl, Tag: Zahl)](/javascript/api/outlook/outlook.seriestime#setenddate-year--month--day-)|Legt den Endtermin einer Terminserie fest.|
||[setstartdate (Datum: Zeichenfolge)](/javascript/api/outlook/outlook.seriestime#setstartdate-date-)|Legt das Startdatum einer Terminserie fest.|
||[setstartdate (Jahr: Zahl, Monat: Zahl, Tag: Zahl)](/javascript/api/outlook/outlook.seriestime#setstartdate-year--month--day-)|Legt das Startdatum einer Terminserie fest.|
||[setstartzeit (Stunden: Zahl, Minuten: Zahl)](/javascript/api/outlook/outlook.seriestime#setstarttime-hours--minutes-)|Legt die Startzeit aller Instanzen einer Terminserie fest, in welcher Zeitzone das Serienmuster festgelegt ist.|
||[setstartzeit (Time: String)](/javascript/api/outlook/outlook.seriestime#setstarttime-time-)|Legt die Startzeit aller Instanzen einer Terminserie fest, in welcher Zeitzone das Serienmuster festgelegt ist.|
