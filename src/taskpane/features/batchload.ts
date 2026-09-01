export interface BatchLoad<T> {
  items: T[];
  queue: (item: T) => void;
  sync: () => Promise<void>;
}

/**
 * Queues loads for every item and syncs once. PowerPoint rejects the whole batch when a
 * single shape refuses a property (Shape.textFrame throws InvalidArgument on shapes that
 * do not support one), so a failed batch is bisected until the offenders are isolated.
 * Returns the items the host refused, in their original order; everything else is loaded.
 */
export async function loadInBatches<T>({ items, queue, sync }: BatchLoad<T>): Promise<T[]> {
  if (items.length === 0) return [];

  items.forEach(queue);
  try {
    await sync();
    return [];
  } catch {
    if (items.length === 1) return items;
  }

  const middle = Math.floor(items.length / 2);
  const head = await loadInBatches({ items: items.slice(0, middle), queue, sync });
  const tail = await loadInBatches({ items: items.slice(middle), queue, sync });
  return head.concat(tail);
}
