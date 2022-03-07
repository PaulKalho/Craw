from datetime import datetime

def sample_responses(input_text):
    user_message = str(input_text).lower()

    if user_message in ("hello", "hallo", "hi"):
        return "Moin"

    if user_message in ("wer bist du"):
        return "Ich bin der Craw Bot"

    if user_message in ("time", "time?"):
        now = datetime.now()
        date_time = now.strftime('%d/%m/%y , %H:%M:%S')

        return str(date_time)


    return "Nix verstehe"