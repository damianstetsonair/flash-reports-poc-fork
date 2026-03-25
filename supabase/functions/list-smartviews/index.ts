import { serve } from "https://deno.land/std@0.168.0/http/server.ts"
import { corsHeaders, handleCors } from "../_shared/cors.ts"

/**
 * Lists all smartviews of type 'project' from AirSaas API.
 * For each smartview, also fetches the project count via /item_ids/.
 * Used by the frontend to let users select which smartview to export.
 */

interface Smartview {
  id: string
  name: string
  type: string
  display: string
  group_by: string
  group_by_level_2: string | null
  description: string | null
  private: boolean
  view_category: string
  created_at: string
  updated_at: string
}

interface AirSaasResponse {
  count: number
  next: string | null
  previous: string | null
  results: Smartview[]
}

interface ItemIdsResponse {
  count: number
  type: string
  item_ids: string[]
}

/** Max concurrent item_ids requests to avoid rate-limiting. */
const CONCURRENCY_LIMIT = 10

/**
 * Fetch project count for a single smartview.
 * Returns 0 on any error (deleted smartview, permissions, etc.).
 */
async function fetchProjectCount(
  baseUrl: string,
  headers: Record<string, string>,
  smartviewId: string,
): Promise<number> {
  try {
    const resp = await fetch(
      `${baseUrl}/smartviews/${smartviewId}/item_ids/`,
      { headers },
    )
    if (!resp.ok) return 0
    const data: ItemIdsResponse = await resp.json()
    return data.item_ids?.length ?? 0
  } catch {
    return 0
  }
}

/**
 * Run promises in batches to respect rate limits.
 */
async function batchedMap<T, R>(
  items: T[],
  fn: (item: T) => Promise<R>,
  batchSize: number,
): Promise<R[]> {
  const results: R[] = []
  for (let i = 0; i < items.length; i += batchSize) {
    const batch = items.slice(i, i + batchSize)
    const batchResults = await Promise.all(batch.map(fn))
    results.push(...batchResults)
  }
  return results
}

serve(async (req) => {
  const corsResponse = handleCors(req)
  if (corsResponse) return corsResponse

  try {
    const apiKey = Deno.env.get('AIRSAAS_API_KEY')
    const baseUrl = Deno.env.get('AIRSAAS_BASE_URL') || 'https://api.airsaas.io/v1'

    if (!apiKey) {
      throw new Error('Missing AIRSAAS_API_KEY environment variable')
    }

    const headers = {
      'Authorization': `Api-Key ${apiKey}`,
      'Content-Type': 'application/json',
    }

    // 1. Fetch all project smartviews (with pagination)
    const allSmartviews: Smartview[] = []
    let url: string | null = `${baseUrl}/smartviews/?type=project&page_size=50`

    while (url) {
      const response = await fetch(url, { headers })

      if (!response.ok) {
        const errorText = await response.text()
        console.error(`AirSaas API error: ${response.status} - ${errorText}`)
        throw new Error(`Failed to fetch smartviews: ${response.status}`)
      }

      const data: AirSaasResponse = await response.json()
      allSmartviews.push(...data.results)

      // Follow pagination
      url = data.next
    }

    console.log(`Fetched ${allSmartviews.length} project smartviews`)

    // 2. Fetch project counts in parallel (batched to avoid rate limits)
    const projectCounts = await batchedMap(
      allSmartviews,
      (sv) => fetchProjectCount(baseUrl, headers, sv.id),
      CONCURRENCY_LIMIT,
    )

    // 3. Pair smartviews with their project counts (frontend handles sorting)
    const indexed = allSmartviews.map((sv, i) => ({ sv, projects_count: projectCounts[i] }))

    return new Response(
      JSON.stringify({
        success: true,
        smartviews: indexed.map(({ sv, projects_count }) => ({
          id: sv.id,
          name: sv.name,
          description: sv.description,
          display: sv.display,
          view_category: sv.view_category,
          private: sv.private,
          updated_at: sv.updated_at,
          projects_count,
        })),
        total: indexed.length,
      }),
      {
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      }
    )
  } catch (error) {
    console.error('List smartviews error:', error)
    return new Response(
      JSON.stringify({
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error',
      }),
      {
        status: 500,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      }
    )
  }
})
