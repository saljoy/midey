import { createFileRoute } from "@tanstack/react-router";
import { getStore } from "@netlify/blobs";

export const Route = createFileRoute("/api/opens")({
  server: {
    handlers: {
      // GET /api/opens — polled by the Open Tracking panel in the app.
      // Returns { opens: { [id]: isoTimestamp } } for every lead that
      // has triggered the tracking pixel at least once.
      GET: async () => {
        try {
          const store = getStore({ name: "outreach-opens" });
          const { blobs } = await store.list();
          const entries: Record<string, string> = {};
          await Promise.all(
            blobs.map(async (b) => {
              try {
                const v = await store.get(b.key);
                if (v) entries[b.key] = v;
              } catch {
                // Skip any individual key that fails to read.
              }
            }),
          );
          return Response.json({ opens: entries });
        } catch (err) {
          console.error("opens fetch error", err);
          // Fail soft — the dashboard just shows 0 opens rather than erroring.
          return Response.json({ opens: {} }, { status: 200 });
        }
      },
    },
  },
});
