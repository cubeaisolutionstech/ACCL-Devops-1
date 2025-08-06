import LoginSystem from './LoginSystem.js';
import Dashboard from './dashboard.js';
import { useState, useEffect } from 'react';

function App() {
  const [user, setUser] = useState(null);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    // Check if user data exists in localStorage
    const userData = localStorage.getItem('user_data');
    if (userData) {
      try {
        const parsedUser = JSON.parse(userData);
        setUser(parsedUser);
      } catch (error) {
        // If there's an error parsing the data, clear it
        localStorage.removeItem('user_data');
        localStorage.removeItem('auth_token');
      }
    }
    setIsLoading(false);
  }, []);

  const handleLogout = () => {
    localStorage.removeItem('user_data');
    localStorage.removeItem('auth_token');
    setUser(null);
  };

  if (isLoading) {
    return <div>Loading...</div>;
  }

  return (
    <div className="App">
      {user && user.user_type === 'admin' ? (
        <Dashboard onLogout={handleLogout} />
      ) : (
        <LoginSystem onLogin={setUser} />
      )}
    </div>
  );
}

export default App;