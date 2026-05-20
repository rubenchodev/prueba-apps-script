function testEmailDelegate(){
  var to_ = "ruben.sanchez@sinova.co,rubenchux@gmail.com";
  var subject_ = "Agrocity: Solicitud de actualización de precios de conceptos - HOY";

  //Referenciamos el contenido de la plantilla
  var htmlBody_ = HtmlService.createHtmlOutputFromFile("Html_MailTemplateUpdateConcepts").getContent();

  //Insertamos la tabla de las especies
  htmlBody_ = htmlBody_.replace("|*CONCEPT_LIST*|", "TABLA");
  
  //Insertamos el vínculo del formulario de actualización
  //htmlBody_ = htmlBody_.replace("|*LINK*|",  webUrl + "?uuid=" + manufacturersDataObj.manufacturerId + "_" + codePack);
  htmlBody_ = htmlBody_.replace("|*LINK*|",  "https://script.google.com/a/macros/hidro.expert/s/AKfycby5t2PiMMKidE2j49LTaf41FeyuVmxLDmbaltLf5NXLQEvX6fI/exec");
  
  sendEmailDelegation_(to_, subject_, htmlBody_);
}

/**
 * Permite enviar un correo a nombre del usuario activo
 */
function sendEmailDelegation_(to_, subject_, htmlBody_) {
  
  // Definimos el usuario activo
  var userDelegateEmail_ = getUserMail();
  //var userDelegateEmail_ = "soporte@hidro.expert";

  // Creamos el servicio para obtener un token asociado al correo de quien abrio el aplicativo
  var service = getEmailService_(userDelegateEmail_);
  //throw service.hasAccess()
  //Se valida que exista un token valido
  if (service.hasAccess()) {

    // Aplicamos el UTF_8 en el asunto por si lleva tildes
    var encodedSubject = Utilities.base64Encode(subject_, Utilities.Charset.UTF_8);
    
    // Creamos estructura Texto con la información del correo
    var emailText_ = "To: " + to_ + "\r\n" +
      //"Subject: " + subject_ + "\r\n" +
      "Subject: =?UTF-8?B?" + encodedSubject + "?=\r\n" + // Para controlar el uso de tíldes
      "Content-Type: text/html; charset=utf-8\r\n" +
      "From: Adquisiciones<" + userDelegateEmail_ + ">\r\n\r\n" +
      htmlBody_;

    // Procesamos los datos para que se permitan enviar correctamente
    emailText_ = Utilities.base64EncodeWebSafe(emailText_, Utilities.Charset.UTF_8);

    // Se define el cuerpo de la petición
    var payload = {
      raw: emailText_
    };

    // Definimos las opciones a enviar
    var options = {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: ("Bearer " + service.getAccessToken()),
      },
      muteHttpExceptions: true,
      payload: JSON.stringify(payload)
    };
    
    // Se realiza la petición
    var response = UrlFetchApp.fetch("https://www.googleapis.com/gmail/v1/users/me/messages/send", options);
    Logger.log(response.getContentText());
    
  } else {

    // Se muestra el error
    console.error(service.getLastError());
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function resetEmailService() {
  getEmailService_().reset();
}

/**
 * Se configura el servicio para delegar el acceso a Gmail
 */
function getEmailService_(userDelegateEmail_) {

  // Definimos las variables a usar
  var provateKey_ = "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQC8kD61h6bQ/3vk\nwPaQkS6IPSqRnpsWDdpBmJH5YL6D7idvKP37PRp7TwwrzC6g6SwAfY9lmd/ukrW+\nYCn8/KX8U1ZGd/5ERJFiHsjximPau7EZOjUvFpYGGRaxRuQSQayCLbdtbUywC5Py\n1CkgmjS4K2Ib0B54/3kioKVwIVYT8LnWMEFVVn4P4edk+83OgAqpkP/YSGEwuY5d\nRpynWg9zFhC9OtifnUXEHB3MzAQsBus4WQwgKwrm8oqJLdiHx5yUIFmngG1irTC8\niPGO00XkbwpFZ1pJzxMkfZ6CRDFeM0UWcWjmdW/jjyewQAKYhmRU/tkd6ZtGw7ZO\nlTfZB2StAgMBAAECggEALFxbBGXnbt/MC+MBnr0pD1AkhaGkpqmXoTr9jIegCnWq\nvVINfl+/rSbJnDwq9BwnB+6CH8r1jbDiJOmGt0fQ3O3WLDnqjbD76vOHHpwKJdlm\nLqMtA+jyXnjkh8w/t7PPm8iZkJrYIcp3gfc7vbnJkVJTgGuT369qliTpBtow3/uo\nCC/PsvxabDIsbOETNNsGcS53BwzPGzBpuk9GkdzzTT9pwS5IxeHjfRfvYtetD71R\nx3ZpoKqx45rjohklPqPVzXzJDuIQXHLwfHHGxaYBdO84yza7f7S3Pa1q3L2Qa2W5\nZYgKBgOum6QXBhR/FgS+miduBy3J69zkxduEXQAL2QKBgQD22qrNbz9On0jN/zut\nY2Ppvs4pHxwpJfiPLHfKYWsPJxj/VBgPEfrKje2aHrNaYOS5h07ilGmf1D8z1R94\nEWBl5VQXzHxAqP2VIVzjzGCrVchhaiXfkfN6MNF43NOACKIyw6G7lAT+0nSSgC7K\nVLDyD4vnBopK9srE68FWOJJxcwKBgQDDjLWBqrbCYDd1Ii4+JAIBbrbhZCFB+KFe\nHsJjEvemKs1MYMzQDyr58zX6azKqqqZepGJgr7bh0dUfsdF7/tYv620CAI5MkDAK\nZeb2BoLS9j+esAbjg+JBrMM3tXB5ZWeMaIFhTuQPtl4/363Gnm2HRujdMYKP8dB8\nLsk3yvTJXwKBgQCKQSQCVPku7pUXEoj47j6KLpxGea8K+MJCZdRdm7wbS3l7nzyu\nTBKbGWxmsiUfTD8uDbNHwXl606PmQsoSNCE6UHuNrLA2gq8rs/jTeKb3rt7AeHPn\nyFWGLQtU1tD511wIdHsDlL5LfdFeL76FAXPszno7lPUlUR11ttickqiqMwKBgDU3\nVrHZ7nCJVK5IxgLdM7qGEbjW0gItEmUVw9rIb6LF5SbS9QnuBskMufTTNW+ijzIU\nVD1xtDhSLUi50IOI/2r9C6ZO1mrZlmlOAAOWFAYrQ3cA+m7yFiG/TUUc2GYBBaP9\nULBVktT/0Thau+TgenglAPwpHBnmOdGBvfqz/dfLAoGBAJx8eVyh1xjVXu1pZyGy\n+eog+ZW1NNlCOQqC/G2dGK/jY0NnAAxqsL05AxN4SYrXpSbqFtGNh8bOM7yBkR3O\nwiHIKGHT/EQIttfipYIkzznvhHeQAMb4Q5ACwgpD7dtUu6Jdm5Agl/WWBT7eugWI\nt4gImMD5JMN1QyIf1DvFtTq7\n-----END PRIVATE KEY-----\n";
  var clientEmail_ = "hicrocotizador-actualizacion-d@project-id-9220749183410745130.iam.gserviceaccount.com";

  return OAuth2.createService('GmailHEXP_PROD:' + userDelegateEmail_)
      // Set the endpoint URL.
      .setTokenUrl('https://oauth2.googleapis.com/token')

      // Set the private key and issuer.
      .setPrivateKey(provateKey_)
      .setIssuer(clientEmail_)

      // Set the name of the user to impersonate. This will only work for
      // Google Apps for Work/EDU accounts whose admin has setup domain-wide
      // delegation:
      // https://developers.google.com/identity/protocols/OAuth2ServiceAccount#delegatingauthority
      .setSubject(userDelegateEmail_)

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getScriptProperties())

      // Set the scope. This must match one of the scopes configured during the
      // setup of domain-wide delegation.
      .setScope('https://www.googleapis.com/auth/gmail.send');
}