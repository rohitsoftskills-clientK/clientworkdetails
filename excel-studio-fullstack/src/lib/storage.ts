import { Product } from '../types/schema';
import { TEMPLATES } from './templates';

const STORAGE_KEY = 'excel_studio_products';
const CORRUPTED_BACKUP_KEY = 'excel_studio_products_corrupted_backup';

/**
 * Structural check so a successfully-parsed but malformed blob (old schema,
 * stray object, etc.) doesn't get handed to the rest of the app as if it
 * were valid.
 */
function isValidProductArray(value: unknown): value is Product[] {
  return (
    Array.isArray(value) &&
    value.every(
      (p) => p && typeof p === 'object' && typeof (p as Product).id === 'string' && Array.isArray((p as Product).sheets)
    )
  );
}

export function getStoredProducts(): Product[] {
  if (typeof window === 'undefined') return TEMPLATES;

  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(TEMPLATES));
    return TEMPLATES;
  }

  try {
    const parsed = JSON.parse(raw);
    if (!isValidProductArray(parsed)) {
      throw new Error('Stored data is not a valid Product[] array');
    }
    return parsed;
  } catch (err) {
    // Don't let a corrupted blob silently vanish: log it clearly and keep a
    // backup copy so it's recoverable, instead of just serving templates and
    // letting the next saveProduct() call overwrite the original data.
    console.error('[storage] Corrupted or invalid product data in localStorage — using default templates for this session.', err);
    try {
      localStorage.setItem(CORRUPTED_BACKUP_KEY, raw);
    } catch {
      // best-effort backup only; ignore quota errors here
    }
    return TEMPLATES;
  }
}

export function saveProduct(product: Product): Product {
  const products = getStoredProducts();
  const index = products.findIndex((p) => p.id === product.id);
  const now = new Date().toISOString();
  const updatedProduct = {
    ...product,
    updated_at: now,
    created_at: product.created_at || now,
  };

  if (index >= 0) {
    products[index] = updatedProduct;
  } else {
    products.unshift(updatedProduct);
  }

  if (typeof window !== 'undefined') {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(products));
    } catch (err) {
      console.error('[storage] Failed to save product — localStorage write failed (quota exceeded or unavailable).', err);
    }
  }
  return updatedProduct;
}

export function getProductById(id: string): Product | undefined {
  const products = getStoredProducts();
  return products.find((p) => p.id === id);
}

export function deleteProduct(id: string): void {
  const products = getStoredProducts().filter((p) => p.id !== id);
  if (typeof window !== 'undefined') {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(products));
    } catch (err) {
      console.error('[storage] Failed to delete product — localStorage write failed.', err);
    }
  }
}

export function createNewProduct(): Product {
  const id = 'prod_' + Math.random().toString(36).substring(2, 9);
  const newProduct: Product = {
    id,
    name: 'Untitled Spreadsheet Product',
    version: '1.0.0',
    author: 'Studio Creator',
    currency: 'USD',
    dateFormat: 'YYYY-MM-DD',
    theme: 'premium',
    sheets: [
      {
        id: 'sheet_1',
        name: 'Sheet 1',
        description: 'Primary data sheet',
        columns: [
          { key: 'item', label: 'Item Name', type: 'text' },
          { key: 'category', label: 'Category', type: 'text' },
          { key: 'amount', label: 'Amount ($)', type: 'currency' },
        ],
        kpis: [
          { label: 'Total Amount', aggregation: 'sum', column: 'amount', format: 'currency' },
          { label: 'Item Count', aggregation: 'count', column: 'item', format: 'number' },
        ],
        rows: [
          ['Consulting Service', 'Services', '4500'],
          ['Software Subscription', 'SaaS', '1200'],
          ['Hardware Purchase', 'Capital Goods', '3800'],
        ],
      },
    ],
  };
  return saveProduct(newProduct);
}
