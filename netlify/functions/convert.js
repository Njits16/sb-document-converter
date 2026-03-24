exports.handler = async function (event) {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: "API-key niet geconfigureerd op de server." }),
    };
  }

  let body;
  try {
    body = JSON.parse(event.body);
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: "Ongeldig verzoek." }) };
  }

  const { text, options } = body;
  if (!text) {
    return { statusCode: 400, body: JSON.stringify({ error: "Geen tekst ontvangen." }) };
  }

  // Bouw de system prompt op basis van gekozen opties
  let systemPrompt =
    "Je bent een document-assistent voor SB Procesmanagement, een freelance interim procesmanager gespecialiseerd in logistiek en warehouse-optimalisatie.\n\n";

  if (options?.style) {
    systemPrompt += `Pas de volgende stijlregels toe:
- Professionele, zakelijke toon — direct en helder
- Gebruik actieve zinsconstructies (vermijd lijdende vorm)
- Structureer het document met duidelijke kopjes waar passend
- Verwijder wollige taal en overbodige herhalingen
- Houd de inhoud bondig maar volledig\n\n`;
  }

  if (options?.improve) {
    systemPrompt += `Verbeter ook:
- Grammatica en spelling
- Leesbaarheid en doorstroming van de tekst
- Logische volgorde van de inhoud\n\n`;
  }

  systemPrompt +=
    "Geef ALLEEN de verbeterde tekst terug. Geen uitleg, geen commentaar, geen inleiding — alleen de inhoud van het document zelf.";

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 4096,
        system: systemPrompt,
        messages: [
          {
            role: "user",
            content: `Verwerk het volgende document:\n\n${text}`,
          },
        ],
      }),
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      return {
        statusCode: response.status,
        body: JSON.stringify({
          error: err.error?.message || `API-fout (${response.status})`,
        }),
      };
    }

    const data = await response.json();
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ result: data.content[0].text }),
    };
  } catch (err) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: "Serverfout: " + err.message }),
    };
  }
};
