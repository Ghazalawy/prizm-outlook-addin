/**
 * Thin ERP API client. Centralises base URL, auth and error handling.
 *
 * Backend endpoints expected (extend on the ERP side as you add views):
 *   POST {apiBase}/outlook/tasks         -> create task from email
 *   POST {apiBase}/outlook/opportunities -> create opportunity
 *   POST {apiBase}/outlook/leads         -> create lead from sender
 *   POST {apiBase}/outlook/tickets       -> create ticket
 *   POST {apiBase}/outlook/link          -> link this email to an existing record
 *   GET  {apiBase}/outlook/lookup?email= -> find ERP records for a contact
 *   GET  {apiBase}/outlook/search?q=&type= -> autocomplete (project/customer/...)
 *   GET  {apiBase}/outlook/refdata       -> priorities, staff, tags, etc.
 *   GET  {apiBase}/outlook/ping          -> health + auth check
 */
import { Config } from './config.js';

class ApiError extends Error {
  constructor(message, status, body) {
    super(message);
    this.status = status;
    this.body = body;
  }
}

async function request(path, { method = 'GET', body, query, signal } = {}) {
  const base = Config.get('apiBase').replace(/\/+$/, '');
  const url = new URL(base + path);
  if (query) {
    Object.entries(query).forEach(([k, v]) => {
      if (v !== undefined && v !== null && v !== '') url.searchParams.set(k, String(v));
    });
  }

  const headers = { 'Accept': 'application/json' };
  const apiKey = Config.get('apiKey');
  if (apiKey) headers['Authorization'] = `Bearer ${apiKey}`;

  let payload;
  if (body !== undefined) {
    headers['Content-Type'] = 'application/json';
    payload = JSON.stringify(body);
  }

  let response;
  try {
    response = await fetch(url.toString(), { method, headers, body: payload, signal });
  } catch (e) {
    throw new ApiError(`Network error: ${e.message}`, 0, null);
  }

  const contentType = response.headers.get('content-type') || '';
  const data = contentType.includes('application/json')
    ? await response.json().catch(() => null)
    : await response.text().catch(() => null);

  if (!response.ok) {
    const msg = (data && (data.message || data.error)) || `${response.status} ${response.statusText}`;
    throw new ApiError(msg, response.status, data);
  }
  return data;
}

export const Api = {
  ping()                        { return request('/outlook/ping'); },
  refdata()                     { return request('/outlook/refdata'); },
  search(type, q)               { return request('/outlook/search', { query: { type, q } }); },
  lookupContact(email)          { return request('/outlook/lookup', { query: { email } }); },
  createTask(payload)           { return request('/outlook/tasks',         { method: 'POST', body: payload }); },
  createOpportunity(payload)    { return request('/outlook/opportunities', { method: 'POST', body: payload }); },
  createLead(payload)           { return request('/outlook/leads',         { method: 'POST', body: payload }); },
  createTicket(payload)         { return request('/outlook/tickets',       { method: 'POST', body: payload }); },
  linkEmail(payload)            { return request('/outlook/link',          { method: 'POST', body: payload }); },
};

export { ApiError };
