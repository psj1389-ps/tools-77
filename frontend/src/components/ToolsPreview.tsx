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
        {/* 섹션 헤더 */}
        <div className="text-center mb-16">
          <h2 className="text-3xl sm:text-4xl font-bold text-gray-900 mb-4">
            인기 도구들
          </h2>
          <p className="text-xl text-gray-600 max-w-3xl mx-auto mb-8">
            가장 많이 사용되는 필수 온라인 도구들을 만나보세요.
          </p>
          
          {/* 카테고리 필터 (데모용) */}
          <div className="flex flex-wrap justify-center gap-3">
            <button className="px-4 py-2 bg-blue-600 text-white rounded-full text-sm font-medium">
              전체
            </button>
            <button className="px-4 py-2 bg-white text-gray-600 hover:bg-gray-100 rounded-full text-sm font-medium transition-colors duration-200">
              변환 도구
            </button>
            <button className="px-4 py-2 bg-white text-gray-600 hover:bg-gray-100 rounded-full text-sm font-medium transition-colors duration-200">
              이미지 도구
            </button>
            <button className="px-4 py-2 bg-white text-gray-600 hover:bg-gray-100 rounded-full text-sm font-medium transition-colors duration-200">
              보안 도구
            </button>
          </div>
        </div>

        {/* 도구 카드 그리드 */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6 mb-12">
          {featuredTools.map((tool) => (
            <Link
              key={tool.id}
              to={tool.path}
              className="group bg-white rounded-xl p-6 shadow-sm hover:shadow-xl transition-all duration-300 transform hover:-translate-y-1 border border-gray-100 hover:border-blue-200"
            >
              {/* 도구 아이콘 */}
              <div className="flex items-center justify-between mb-4">
                <div className="p-3 bg-blue-50 text-blue-600 rounded-lg group-hover:bg-blue-600 group-hover:text-white transition-colors duration-300">
                  {getIcon(tool.icon)}
                </div>
                
                {/* 인기 배지 */}
                <div className="flex items-center space-x-1 text-blue-500">
                  <Star className="w-4 h-4 fill-current" />
                  <span className="text-xs font-medium text-gray-600">인기</span>
                </div>
              </div>

              {/* 도구 정보 */}
              <h3 className="text-lg font-semibold text-gray-900 mb-2 group-hover:text-blue-600 transition-colors duration-200">
                {tool.name}
              </h3>
              
              <p className="text-gray-600 text-sm mb-4 line-clamp-2">
                {tool.description}
              </p>

              {/* 카테고리 태그 */}
              <div className="flex items-center justify-between">
                <span className={`px-2 py-1 rounded-full text-xs font-medium ${getCategoryColor(tool.category)}`}>
                  {tool.category === 'converter' && '변환'}
                  {tool.category === 'image' && '이미지'}
                  {tool.category === 'text' && '텍스트'}
                  {tool.category === 'security' && '보안'}
                  {tool.category === 'generator' && '생성'}
                  {tool.category === 'design' && '디자인'}
                </span>
                
                <ArrowRight className="w-4 h-4 text-gray-400 group-hover:text-blue-600 group-hover:translate-x-1 transition-all duration-200" />
              </div>
            </Link>
          ))}
        </div>

        {/* 더 많은 도구 보기 버튼 */}
        <div className="text-center">
          <Link
            to="/tools"
            className="inline-flex items-center px-8 py-4 bg-white text-blue-600 border-2 border-blue-600 rounded-lg font-semibold hover:bg-blue-600 hover:text-white transition-all duration-200 transform hover:scale-105"
          >
            모든 도구 보기
            <ArrowRight className="ml-2 w-5 h-5" />
          </Link>
        </div>
      </div>
    </section>
  );
};

export default ToolsPreview;