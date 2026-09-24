import React, { useState } from 'react';
import { AuthProvider, useAuth } from './context/AuthContext';
import { Sidebar } from './components/Sidebar';
import { Navbar } from './components/Navbar';
import { DocumentViewerModal } from './components/DocumentViewerModal';
import { CheckListManager } from './components/CheckListManager';
import { DOCUMENTS, INITIAL_MENU_ITEMS } from './data/mockData';
import { Search, Eye, FileText, Lock } from 'lucide-react';

const DashboardContent = () => {
  const { user } = useAuth();
  const [activeCategory, setActiveCategory] = useState('evaluacion-riesgos');
  const [menuItems, setMenuItems] = useState(INITIAL_MENU_ITEMS);
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedDoc, setSelectedDoc] = useState(null);

  const filteredDocuments = DOCUMENTS.filter(doc => {
    if (doc.category !== activeCategory) return false;
    if (searchTerm && !doc.title.toLowerCase().includes(searchTerm.toLowerCase())) return false;

    if (user.role === 'Centro') return doc.centerId === user.centerId;
    if (user.role === 'Zona') return doc.zoneId === user.zoneId;
    if (user.role === 'Global') return true;

    return false;
  });

  return (
    <div className="flex min-h-screen bg-slate-100 font-sans">
      <Sidebar 
        activeCategory={activeCategory} 
        setActiveCategory={setActiveCategory}
        menuItems={menuItems}
        setMenuItems={setMenuItems}
      />

      <div className="flex-1 flex flex-col">
        <Navbar />

        <main className="p-8 flex-1">
          {activeCategory === 'menu-checks' ? (
            <CheckListManager />
          ) : (
            <div className="space-y-6">
              
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white p-4 rounded-xl border border-slate-200">
                <div className="relative flex-1 max-w-md">
                  <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input 
                    type="text" 
                    placeholder="Buscar documento por título..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    className="w-full pl-9 pr-4 py-2 text-xs border border-slate-300 rounded-lg focus:outline-none"
                  />
                </div>

                <div className="text-xs text-slate-500 font-medium">
                  Documentos permitidos ({user.role}): <span className="font-bold text-slate-800">{filteredDocuments.length}</span>
                </div>
              </div>

              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="px-6 py-4 border-b border-slate-200 bg-slate-50">
                  <h3 className="font-bold text-slate-800 text-sm">
                    {menuItems.find(m => m.id === activeCategory)?.label || 'Documentos'}
                  </h3>
                </div>

                {filteredDocuments.length === 0 ? (
                  <div className="p-12 text-center">
                    <Lock className="w-8 h-8 text-slate-300 mx-auto mb-3" />
                    <p className="text-sm font-semibold text-slate-600">Sin acceso a esta carpeta</p>
                    <p className="text-xs text-slate-400 mt-1">Tu nivel ({user.role}) no tiene permisos para ver documentos aquí.</p>
                  </div>
                ) : (
                  <div className="divide-y divide-slate-100">
                    {filteredDocuments.map((doc) => (
                      <div key={doc.id} className="p-4 px-6 flex items-center justify-between hover:bg-slate-50">
                        <div className="flex items-center gap-4">
                          <div className="p-2.5 bg-blue-50 text-blue-600 rounded-lg">
                            <FileText className="w-5 h-5" />
                          </div>
                          <div>
                            <h4 className="font-bold text-slate-800 text-sm">{doc.title}</h4>
                            <p className="text-xs text-slate-400 mt-0.5">{doc.centerName} • {doc.zoneName}</p>
                          </div>
                        </div>

                        <button 
                          onClick={() => setSelectedDoc(doc)}
                          className="flex items-center gap-2 px-3 py-1.5 bg-slate-900 hover:bg-blue-600 text-white rounded-lg text-xs font-semibold"
                        >
                          <Eye className="w-3.5 h-3.5" /> Abrir Visor Protegido
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>

            </div>
          )}
        </main>
      </div>

      <DocumentViewerModal 
        document={selectedDoc} 
        onClose={() => setSelectedDoc(null)} 
      />
    </div>
  );
};

export default function App() {
  return (
    <AuthProvider>
      <DashboardContent />
    </AuthProvider>
  );
}
