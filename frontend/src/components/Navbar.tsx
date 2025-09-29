import { useState } from 'react';
import { Link } from 'react-router-dom';
import { Menu, X, Search, Settings } from 'lucide-react';

const Navbar = () => {
  return (
    <nav className="bg-white/80 backdrop-blur-md shadow-sm border-b border-gray-100 sticky top-0 z-50">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex justify-between items-center h-16">
          {/* 로고 */}
          <Link to="/" className="flex items-center space-x-2">
            <div className="w-12 h-12 rounded-full overflow-hidden">
              <img 
                src="https://trae-api-sg.mchost.guru/api/ide/v1/text_to_image?prompt=circular%20purple%20logo%20with%20white%20text%20%2277%20TOOLS%22%20modern%20design%20clean%20typography&image_size=square" 
                alt="77 TOOLS Logo" 
                className="w-full h-full object-cover"
              />
            </div>
            <span className="text-xl font-bold text-gray-900">
              77+ Popular Tools
            </span>
          </Link>

          {/* 검색창 */}
          <div className="flex-1 flex justify-center px-2 lg:ml-6 lg:justify-end">
            <div className="max-w-md w-full lg:max-w-xs">
              <label htmlFor="search" className="sr-only">Search</label>
              <div className="relative">
                <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                  <Search className="h-5 w-5 text-gray-400" aria-hidden="true" />
                </div>
                <input
                  id="search"
                  name="search"
                  className="block w-full pl-10 pr-3 py-2 border border-gray-300 rounded-md leading-5 bg-white placeholder-gray-500 focus:outline-none focus:placeholder-gray-400 focus:ring-1 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                  placeholder="Search"
                  type="search"
                />
              </div>
            </div>
          </div>

        </div>
      </div>
    </nav>
  );
};

export default Navbar;