import React from 'react';
import Navbar from '../components/Navbar';
import Hero from '../components/Hero';
import PopularTools from '../components/PopularTools';
import ToolsPreview from '../components/ToolsPreview';
import Footer from '../components/Footer';

const Home: React.FC = () => {
  return (
    <div className="min-h-screen bg-gray-50">
      <Navbar />
      <main>
        <Hero />
        <PopularTools />
        <ToolsPreview />
      </main>
      <Footer />
    </div>
  );
};

export default Home;