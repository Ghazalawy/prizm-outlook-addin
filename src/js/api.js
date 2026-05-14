/**
 * Thin ERP API client. Centralises base URL, auth and error handling.
 *
 * Backend lives in the Perfex `outlookapi` module, controller `Bridge`:
 *   POST {apiBase}/tasks         -> create task from email
 *   POST {apiBase}/opportunities -> create opportunity
 *   POST {apiBase}/leads         -> create lead from sender
 *   POST {apiBase}/tickets       -> create ticket
 *   POST {apiBase}/link          -> link this email to an existing record
 *   GET  {apiBase}/lookup?email= -> find ERP records for a contact
 *   GET  {apiBase}/search?type=&q= -> autocomplete (project/customer/...)
 *   GET  {apiBase}/refdata       -> priorities, staff, tags
 *   GET  {apiBase}/ping          -> health + auth check
 *
 * apiBase default points to `outlookapi/bridge` on the Perfex install.
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
  ping()                        { return request('/ping'); },
  refdata()                     { return request('/refdata'); },
  search(type, q, limit)        { return request('/search', { query: { type, q, limit } }); },
  lookupContact(email)          { return request('/lookup', { query: { email } }); },
  emailStatus(messageId)        { return request('/email_status', { query: { messageId } }); },
  createTask(payload)           { return request('/tasks',         { method: 'POST', body: payload }); },
  createOpportunity(payload)    { return request('/opportunities', { method: 'POST', body: payload }); },
  createLead(payload)           { return request('/leads',         { method: 'POST', body: payload }); },
  createTicket(payload)         { return request('/tickets',       { method: 'POST', body: payload }); },
  linkEmail(payload)            { return request('/link',          { method: 'POST', body: payload }); },
};

export { ApiError };
