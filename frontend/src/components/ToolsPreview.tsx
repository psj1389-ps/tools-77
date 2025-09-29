import { Link } from 'react-router-dom';
import { FileText, Image, Type, Shield, QrCode, Palette, ArrowRight, Star } from 'lucide-react';
import { TOOLS } from '../data/constants';

const ToolsPreview = () => {
  const getIcon = (iconName: string) => {
    switch (iconName) {
      case 'FileText':
        return <FileText className="w-6 h-6" />;
      case 'Image':
        return <Image className="w-6 h-6" />;
      case 'Type':
        return <Type className="w-6 h-6" />;
      case 'Shield':
        return <Shield className="w-6 h-6" />;
      case 'QrCode':
        return <QrCode className="w-6 h-6" />;
      case 'Palette':
        return <Palette className="w-6 h-6" />;
      default:
        return <FileText className="w-6 h-6" />;
    }
  };

  const getCategoryColor = (category: string) => {
    switch (category) {
      case 'converter':
        return 'bg-blue-100 text-blue-700';
      case 'image':
        return 'bg-green-100 text-green-700';
      case 'text':
        return 'bg-purple-100 text-purple-700';
      case 'security':
        return 'bg-blue-100 text-blue-700';
      case 'generator':
        return 'bg-blue-100 text-blue-700';
      case 'design':
        return 'bg-pink-100 text-pink-700';
      default:
        return 'bg-gray-100 text-gray-700';
    }
  };

  const featuredTools = TOOLS.filter(tool => tool.featured);

  return (
    <section className="py-20 bg-gray-50">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        {/* 도구 카드 그리드 */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
          {featuredTools.map((tool) => (
            <Link
              key={tool.id}
              to={tool.path}
              className="group bg-white rounded-xl p-4 shadow-sm hover:shadow-xl transition-all duration-300 transform hover:-translate-y-1 border border-gray-100 hover:border-blue-200"
            >
              {/* 도구 아이콘 */}
              <div className="flex items-center justify-between mb-2">
                <div className="p-2 bg-blue-50 text-blue-600 rounded-lg group-hover:bg-blue-600 group-hover:text-white transition-colors duration-300">
                  {getIcon(tool.icon)}
                </div>
                
                {/* 인기 배지 */}
                <div className="flex items-center space-x-1 text-blue-500">
                  <Star className="w-3 h-3 fill-current" />
                  <span className="text-xs font-medium text-gray-600">인기</span>
                </div>
              </div>

              {/* 도구 정보 */}
              <h3 className="text-base font-semibold text-gray-900 mb-1 group-hover:text-blue-600 transition-colors duration-200">
                {tool.name}
              </h3>
              
              <p className="text-gray-600 text-xs mb-2 line-clamp-2">
                {tool.description}
              </p>

              {/* 카테고리 태그 */}
              <div className="flex items-center justify-between">
                <span className={`px-2 py-0.5 rounded-full text-xs font-medium ${getCategoryColor(tool.category)}`}>
                  {tool.category === 'converter' && '변환'}
                  {tool.category === 'image' && '이미지'}
                  {tool.category === 'text' && '텍스트'}
                  {tool.category === 'security' && '보안'}
                  {tool.category === 'generator' && '생성'}
                  {tool.category === 'design' && '디자인'}
                </span>
                
                <ArrowRight className="w-3 h-3 text-gray-400 group-hover:text-blue-600 group-hover:translate-x-1 transition-all duration-200" />
              </div>
            </Link>
          ))}
        </div>
      </div>
    </section>
  );
};

export default ToolsPreview;