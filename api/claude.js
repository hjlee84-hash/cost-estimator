export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();

  const { imageData, mimeType, textPrompt } = req.body;
  const apiKey = process.env.GEMINI_API_KEY;
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;

  let parts = [];

  if (imageData) {
    parts = [
      { inline_data: { mime_type: mimeType, data: imageData } },
      { text: "이 BOM 이미지에서 부품 정보를 추출해줘. W, L, T 값이 있는 부품만 추출하고, 없으면 null로. 반드시 아래 JSON 배열만 응답해. 다른 텍스트나 마크다운 없이 순수 JSON만.\n[{\"no\":1,\"name\":\"MTG PLATE\",\"material\":\"AA3003 H16\",\"w\":159,\"l\":109,\"t\":4,\"qty\":1,\"weight\":\"143.0g\",\"remark\":\"\"},...]" }
    ];
  } else {
    parts = [
      { text: `아래 BOM 텍스트에서 부품 정보를 추출해줘. W,L,T 있는 부품만. 반드시 JSON 배열만 응답해. 다른 텍스트나 마크다운 없이 순수 JSON만.\n[{"no":1,"name":"MTG PLATE","material":"AA3003 H16","w":159,"l":109,"t":4,"qty":1,"weight":"143.0g","remark":""},...]\n\nBOM:\n${textPrompt}` }
    ];
  }

  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ contents: [{ parts }] }),
  });

  const data = await response.json();
  res.status(response.status).json(data);
}
