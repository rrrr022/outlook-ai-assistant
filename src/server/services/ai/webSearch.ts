import axios from 'axios';

type SearchResult = {
  title: string;
  link: string;
  snippet?: string;
  source?: string;
};

const SERPAPI_ENDPOINT = 'https://serpapi.com/search.json';

const isNewsQuery = (query: string): boolean => {
  return /(news|headline|breaking|latest|press|announcement|update)/i.test(query);
};

const isRealtimeQuery = (query: string): boolean => {
  return /(weather|forecast|temperature|news|headline|latest|breaking|stock|price|quote|score|sports|traffic|right now|currently|today|this morning|this afternoon|this evening)/i.test(query);
};

export const needsWebSearch = (query: string): boolean => {
  return isRealtimeQuery(query);
};

export async function runWebSearch(query: string, maxResults = 5): Promise<SearchResult[]> {
  const apiKey = process.env.SERPAPI_KEY;
  if (!apiKey) {
    return [];
  }

  const engine = isNewsQuery(query) ? 'google_news' : 'google';

  const response = await axios.get(SERPAPI_ENDPOINT, {
    params: {
      q: query,
      engine,
      api_key: apiKey,
      num: Math.max(1, Math.min(10, maxResults)),
    },
    timeout: 10000,
  });

  const data = response.data || {};

  const organicResults: SearchResult[] = (data.organic_results || []).map((item: any) => ({
    title: item.title,
    link: item.link,
    snippet: item.snippet || item.description,
    source: item.source,
  }));

  const newsResults: SearchResult[] = (data.news_results || []).map((item: any) => ({
    title: item.title,
    link: item.link,
    snippet: item.snippet || item.description,
    source: item.source,
  }));

  const combined = [...newsResults, ...organicResults].filter(r => r.title && r.link);
  return combined.slice(0, maxResults);
}

export function formatSearchResults(results: SearchResult[]): string {
  if (!results.length) {
    return '';
  }

  const lines = results.map((r, index) => {
    const snippet = r.snippet ? ` - ${r.snippet}` : '';
    const source = r.source ? ` (${r.source})` : '';
    return `${index + 1}. ${r.title}${source} — ${r.link}${snippet}`;
  });

  return `\n\nWeb search results (most recent):\n${lines.join('\n')}`;
}
