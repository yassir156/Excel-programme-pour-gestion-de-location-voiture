import axios from 'axios';

const baseURL = window.carRentalApp?.apiBaseUrl || 'http://127.0.0.1:4310/api';

const api = axios.create({ baseURL });

api.interceptors.request.use((config) => {
  const token = localStorage.getItem('car_rental_token');
  if (token) {
    config.headers.Authorization = `Bearer ${token}`;
  }
  return config;
});

api.interceptors.response.use(
  (response) => response,
  (error) => {
    if (error.response?.status === 401) {
      localStorage.removeItem('car_rental_token');
      localStorage.removeItem('car_rental_user');
      if (!window.location.hash.includes('/login')) {
        window.location.hash = '#/login';
      }
    }
    return Promise.reject(error);
  }
);

export function getErrorMessage(err, fallback = 'Une erreur est survenue.') {
  return err?.response?.data?.message || fallback;
}

export const uploadsBaseUrl = baseURL.replace(/\/api$/, '');

export default api;
