import axios from 'axios';

// Create an axios instance
const api = axios.create({
  baseURL: process.env.REACT_APP_API_URL || 'http://localhost:5000/api',
  headers: {
    'Content-Type': 'application/json'
  }
});

// Request interceptor for adding the auth token
api.interceptors.request.use(
  (config) => {
    // You can add common headers or modify request config here
    return config;
  },
  (error) => {
    return Promise.reject(error);
  }
);

// Response interceptor for handling common errors
api.interceptors.response.use(
  (response) => {
    return response;
  },
  (error) => {
    // Handle common errors
    if (error.response) {
      switch (error.response.status) {
        case 401:
          // Unauthorized - Clear token if we get a 401 response
          sessionStorage.removeItem('auth_token');
          break;
        case 403:
          // Forbidden
          console.error('Access forbidden');
          break;
        case 500:
          // Server error
          console.error('Server error occurred');
          break;
        default:
          break;
      }
    }
    
    return Promise.reject(error);
  }
);

// Helper methods for the API service
const apiService = {
  // Set auth token in headers
  setAuthToken: (token) => {
    if (token) {
      api.defaults.headers.common['Authorization'] = `Bearer ${token}`;
    }
  },
  
  // Remove auth token from headers
  removeAuthToken: () => {
    delete api.defaults.headers.common['Authorization'];
  },
  
  // GET request
  get: async (url, config = {}) => {
    return api.get(url, config);
  },
  
  // POST request
  post: async (url, data = {}, config = {}) => {
    return api.post(url, data, config);
  },
  
  // PUT request
  put: async (url, data = {}, config = {}) => {
    return api.put(url, data, config);
  },
  
  // DELETE request
  delete: async (url, config = {}) => {
    return api.delete(url, config);
  }
};

export default apiService;