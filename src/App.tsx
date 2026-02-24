/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import {StrictMode} from 'react';
import ProductionPlanMaker from './components/ProductionPlanMaker';

export default function App() {
  return (
    <div className="min-h-screen bg-gray-50 py-12">
      <ProductionPlanMaker />
    </div>
  );
}
