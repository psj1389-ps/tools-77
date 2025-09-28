import { useState } from 'react'
import { Upload, FileText, Download, AlertCircle } from 'lucide-react'
import { Button } from './components/ui/button'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from './components/ui/card'
import { Alert, AlertDescription } from './components/ui/alert'
import { toast } from 'sonner'

import { BrowserRouter as Router, Routes, Route } from "react-router-dom";
import Home from "@/pages/Home";

function App() {
  const [file, setFile] = useState<File | null>(null)
  const [isConverting, setIsConverting] = useState(false)
  const [error, setError] = useState<string | null>(null)

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-4">
      <div className="max-w-4xl mx-auto">
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-gray-900 mb-2">
            PDF to DOCX Converter
          </h1>
          <p className="text-lg text-gray-600">
            Convert your PDF files to editable DOCX format instantly
          </p>
          <p className="text-sm text-green-600 mt-2">
            ✅ Connected to GitHub - Auto-deployment enabled!
          </p>
        </div>
      </div>
    </div>
  )
}

export default App
