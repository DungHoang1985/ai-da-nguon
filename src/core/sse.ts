/** Bộ tách luồng Server-Sent Events: nhận từng mẩu văn bản, trả ra nội dung từng dòng "data:". */
export function createSSEParser(onData: (data: string) => void) {
  let buffer = "";
  return {
    push(chunk: string) {
      buffer += chunk;
      let idx: number;
      while ((idx = buffer.indexOf("\n")) >= 0) {
        const line = buffer.slice(0, idx).replace(/\r$/, "");
        buffer = buffer.slice(idx + 1);
        if (line.startsWith("data:")) {
          const data = line.slice(5).trimStart();
          if (data && data !== "[DONE]") onData(data);
        }
      }
    },
    end() {
      const line = buffer.replace(/\r$/, "");
      buffer = "";
      if (line.startsWith("data:")) {
        const data = line.slice(5).trimStart();
        if (data && data !== "[DONE]") onData(data);
      }
    },
  };
}

export async function readSSE(body: ReadableStream<Uint8Array>, onData: (data: string) => void): Promise<void> {
  const reader = body.getReader();
  const decoder = new TextDecoder();
  const parser = createSSEParser(onData);
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    parser.push(decoder.decode(value, { stream: true }));
  }
  parser.end();
}
