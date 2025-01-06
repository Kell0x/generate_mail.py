# Importation des classes nécessaires depuis la bibliothèque independentsoft.msg
from independentsoft.msg import Message
from independentsoft.msg import Recipient
from independentsoft.msg import ObjectType
from independentsoft.msg import DisplayType
from independentsoft.msg import RecipientType
from independentsoft.msg import MessageFlag
from independentsoft.msg import StoreSupportMask

# Importation des bibliothèques standard
import re  # Pour utiliser les expressions régulières
from datetime import datetime  # Pour manipuler les dates et heures

# Fonction pour encoder les caractères spéciaux en entités HTML
def encode_html_characters(unencoded_str):
    """
    Converts a string with special characters into an HTML-compatible version.

    Args:
        unencoded_str (str): The raw text to encode.

    Returns:
        str: A string encoded with special characters replaced by HTML entities,
             and newlines replaced with <br>.
    """
    # Convertit les caractères non-ASCII en entités XML (par ex. é → &#233;)
    # Remplace aussi les sauts de ligne par <br> pour compatibilité HTML
    return unencoded_str.encode('ascii', 'xmlcharrefreplace').decode().replace("\n", "<br>")

# Fonction pour valider si une chaîne est une adresse e-mail valide
def validate_mail_address(am_i_a_mail_address):
    """
    Validates that a given string follows the standard email address format.

    Args:
        am_i_a_mail_address (str): The string to validate.

    Raises:
        AssertionError: If the email address is not valid.
    """
    # Expression régulière pour vérifier le format d'une adresse e-mail standard
    email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    # Utilise assert pour lever une erreur si le format n'est pas valide
    assert re.match(email_regex, am_i_a_mail_address), f"{am_i_a_mail_address} is not a valid mail address."

# Fonction pour préparer un destinataire en fonction de son adresse e-mail et de son type (TO ou CC)
def prepare_recipients(recipient_mail, recipient_type=RecipientType.TO):
    """
    Creates and configures a Recipient object for an email recipient.

    Args:
        recipient_mail (str): The recipient's email address.
        recipient_type (RecipientType): The type of recipient (TO or CC, default is TO).

    Returns:
        Recipient: A configured Recipient object.
    """
    # Vérifie que l'adresse e-mail est valide
    validate_mail_address(recipient_mail)
    # Crée un nouvel objet Recipient et configure ses propriétés
    recipient = Recipient()
    recipient.address_type = "SMTP"  # Type d'adresse (protocole utilisé)
    recipient.display_type = DisplayType.MAIL_USER  # Type d'utilisateur affiché
    recipient.object_type = ObjectType.MAIL_USER  # Type d'objet utilisateur
    recipient.display_name = recipient_mail
    recipient.email_address = recipient_mail  # Définit l'adresse e-mail
    recipient.recipient_type = recipient_type  # Définit si le destinataire est TO ou CC
    return recipient  # Retourne l'objet configuré

# Fonction pour générer un nom de fichier pour sauvegarder un message .msg
def generate_msg_filename(recipients, subject="GenericMessage"):
    """
    Generates a unique and descriptive filename for a .msg email file.

    Args:
        recipients (str or list): Email address or list of email addresses for the main recipients.
        subject (str): The subject of the message (default is "GenericMessage").

    Returns:
        str: A filename in the format `YYYYMMDD_HHMMSS_FirstRecipient_Subject.msg`.
    """
    # Génère un timestamp au format YYYYMMDD_HHMMSS
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Formate la partie des destinataires dans le nom du fichier
    if isinstance(recipients, list):  # Si plusieurs destinataires
        first_recipient = recipients[0].split("@")[0]  # Récupère le nom avant le '@' du premier
        if len(recipients) > 1:
            recipients_str = f"{first_recipient}+others"  # Ajoute '+others' s'il y en a plusieurs
        else:
            recipients_str = first_recipient  # Utilise juste le premier si seul
    else:  # Si un seul destinataire
        recipients_str = recipients.split("@")[0] if "@" in recipients else recipients
    
    # Nettoie et limite le sujet à 30 caractères, remplaçant les espaces par des underscores
    subject_clean = subject.replace(" ", "_")[:30]
    
    # Combine les différentes parties pour générer un nom de fichier unique
    filename = f"{timestamp}_{recipients_str}_{subject_clean}.msg"
    return filename

# Fonction principale pour générer un message e-mail au format .msg
def generate_mail(to, cc, subject, body):
    """
    Creates a .msg email file with the provided information.

    Args:
        to (str or list): Email address(es) for the primary recipients (TO).
        cc (str or list): Email address(es) for the carbon copy recipients (CC).
        subject (str): The subject of the email.
        body (str): The body of the email in plain text.

    Actions:
        - Saves a .msg file with an automatically generated name.
        - Adds TO and CC recipients, encodes the email body in HTML and RTF.

    Returns:
        None.
    """
    # Crée un nouvel objet Message
    message = Message()

    # Prépare la liste des destinataires principaux (TO)
    recipient_to = []
    display_to = ""
    if isinstance(to, str):  # Si un seul destinataire
        recipient_to.append(prepare_recipients(to, RecipientType.TO))
        display_to = to
    elif isinstance(to, list):  # Si plusieurs destinataires
        for item in to:
            recipient_to.append(prepare_recipients(item, RecipientType.TO))
        display_to = "; ".join(to)
    else:  # Type invalide
        raise TypeError("The variable 'to' must be an email address (str) or a list of email addresses (list).")

    # Prépare la liste des destinataires en copie (CC)
    recipient_cc = []
    display_cc = ""
    if cc != "":  # Si un ou plusieurs destinataires en copie sont spécifiés
        if isinstance(cc, str):  # Si un seul destinataire en copie
            recipient_cc.append(prepare_recipients(cc, RecipientType.CC))
            display_cc = cc
        elif isinstance(cc, list):  # Si plusieurs destinataires en copie
            for item in cc:
                recipient_cc.append(prepare_recipients(item, RecipientType.CC))
            display_cc = "; ".join(cc)
        else:  # Type invalide
            raise TypeError("The variable 'cc' must be an email address (str) or a list of email addresses (list).")

    # Encode le corps du message en HTML
    html_body = encode_html_characters(body)
    # Ajoute une couche de compatibilité RTF au corps HTML
    html_body_with_rtf = "{\\rtf1\\ansi\\ansicpg1252\\fromhtml1 \\htmlrtf0 " + html_body + "}"
    rtf_body = html_body_with_rtf.encode("utf-8")  # Encode en UTF-8 pour le format RTF

    # Configure les propriétés du message
    message.subject = subject  # Définit le sujet du message
    message.body_html_text = html_body  # Définit le corps HTML
    message.body_rtf = rtf_body  # Définit le corps RTF
    message.display_to = display_to
    message.display_cc = display_cc
    for recipient in recipient_to:  # Ajoute les destinataires principaux
        message.recipients.append(recipient)
    for recipient in recipient_cc:  # Ajoute les destinataires en copie
        message.recipients.append(recipient)
    message.message_flags.append(MessageFlag.UNSENT)  # Indique que le message n'a pas été envoyé
    message.store_support_masks.append(StoreSupportMask.CREATE)  # Précise que le fichier doit être créé

    # Génère un nom de fichier pour sauvegarder le message
    filename = generate_msg_filename(to, subject)
    # Sauvegarde le message au format .msg
    message.save(filename)
