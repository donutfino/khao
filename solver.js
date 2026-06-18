// อัลกอริทึมสำหรับคำนวณแบ่งกลุ่มพิกัดด้วย K-Means และจัดลำดับเส้นทางด้วย 2-opt TSP
window.RoutingSolver = (function() {

  // คำนวณระยะห่างทางภูมิศาสตร์ด้วยสูตร Haversine (กิโลเมตร)
  function getDistance(lat1, lng1, lat2, lng2) {
    const R = 6371; // รัศมีโลก
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLng = (lng2 - lng1) * Math.PI / 180;
    const a = 
      Math.sin(dLat / 2) * Math.sin(dLat / 2) +
      Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * 
      Math.sin(dLng / 2) * Math.sin(dLng / 2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return R * c;
  }

  // คำนวณระยะทางรวมของเส้นทาง
  function calculateRouteDistance(route, depot, isRoundTrip) {
    if (route.length === 0) return 0;
    let distance = 0;
    
    // ระยะทางจากคลังไปยังจุดส่งจุดแรก
    distance += getDistance(depot.lat, depot.lng, route[0].lat, route[0].lng);
    
    // ระยะทางระหว่างจุดจอดส่งแต่ละจุด
    for (let i = 0; i < route.length - 1; i++) {
      distance += getDistance(route[i].lat, route[i].lng, route[i + 1].lat, route[i + 1].lng);
    }
    
    // ระยะทางขากลับคลัง (กรณีเป็น Round Trip)
    if (isRoundTrip) {
      distance += getDistance(route[route.length - 1].lat, route[route.length - 1].lng, depot.lat, depot.lng);
    }
    
    return distance;
  }

  // ปรับปรุงลำดับเส้นทางให้สั้นที่สุดด้วย 2-opt Local Search
  function optimize2Opt(route, depot, isRoundTrip) {
    let improved = true;
    let bestRoute = [...route];
    let bestDist = calculateRouteDistance(bestRoute, depot, isRoundTrip);
    let iterations = 0;
    const maxIterations = 200;

    while (improved && iterations < maxIterations) {
      improved = false;
      iterations++;
      
      for (let i = 0; i < bestRoute.length - 1; i++) {
        for (let j = i + 1; j < bestRoute.length; j++) {
          const newRoute = [...bestRoute];
          // สลับจุดจอดระหว่างดัชนี i ถึง j
          let left = i;
          let right = j;
          while (left < right) {
            const temp = newRoute[left];
            newRoute[left] = newRoute[right];
            newRoute[right] = temp;
            left++;
            right--;
          }

          const newDist = calculateRouteDistance(newRoute, depot, isRoundTrip);
          if (newDist < bestDist - 0.0001) {
            bestRoute = newRoute;
            bestDist = newDist;
            improved = true;
          }
        }
      }
    }
    return bestRoute;
  }

  /**
   * แบ่งกลุ่มลูกค้าออกเป็น K กลุ่มด้วยวิธี K-Means Clustering
   */
  function runKMeans(customers, k, maxIterations = 20) {
    if (customers.length === 0) return [];
    if (k >= customers.length) {
      return customers.map(c => [c]);
    }

    // 1. กำหนดจุดศูนย์กลางกลุ่มเริ่มต้น (Centroids Initialization) ด้วย K-Means++
    let centroids = [];
    // เลือกจุดแรกแบบสุ่ม
    centroids.push({
      lat: customers[0].lat,
      lng: customers[0].lng
    });

    // เลือกจุดถัดๆ ไป โดยหาจุดที่อยู่ห่างจากเซนทรอยด์เดิมมากที่สุด
    for (let cIdx = 1; cIdx < k; cIdx++) {
      let maxDist = -1;
      let nextCentroid = null;
      
      customers.forEach(cust => {
        // หาระยะทางที่ใกล้ที่สุดไปยังเซนทรอยด์ที่เลือกไปแล้ว
        let minDistToExisting = Infinity;
        centroids.forEach(ctr => {
          let dist = getDistance(cust.lat, cust.lng, ctr.lat, ctr.lng);
          if (dist < minDistToExisting) {
            minDistToExisting = dist;
          }
        });

        if (minDistToExisting > maxDist) {
          maxDist = minDistToExisting;
          nextCentroid = { lat: cust.lat, lng: cust.lng };
        }
      });

      if (nextCentroid) {
        centroids.push(nextCentroid);
      } else {
        // กันพลาดกรณีหาไม่ได้ ให้หยิบแบบสุ่ม
        let randCust = customers[Math.floor(Math.random() * customers.length)];
        centroids.push({ lat: randCust.lat, lng: randCust.lng });
      }
    }

    let assignments = new Array(customers.length);
    let centroidsChanged = true;
    let iterations = 0;

    // 2. ลูปคำนวณและปรับเปลี่ยนจุดศูนย์กลางกลุ่ม
    while (centroidsChanged && iterations < maxIterations) {
      centroidsChanged = false;
      iterations++;

      // ก) Assign: จัดสรรลูกค้าเข้าเซนทรอยด์ที่ใกล้ที่สุด
      for (let i = 0; i < customers.length; i++) {
        let minDist = Infinity;
        let bestCluster = 0;
        
        for (let j = 0; j < k; j++) {
          let dist = getDistance(customers[i].lat, customers[i].lng, centroids[j].lat, centroids[j].lng);
          if (dist < minDist) {
            minDist = dist;
            bestCluster = j;
          }
        }

        if (assignments[i] !== bestCluster) {
          assignments[i] = bestCluster;
          centroidsChanged = true;
        }
      }

      // ข) Update: คำนวณเซนทรอยด์ใหม่จากค่าเฉลี่ยของพิกัดในกลุ่ม
      let clusterSums = Array.from({ length: k }, () => ({ lat: 0, lng: 0, count: 0 }));
      for (let i = 0; i < customers.length; i++) {
        let clusterIdx = assignments[i];
        clusterSums[clusterIdx].lat += customers[i].lat;
        clusterSums[clusterIdx].lng += customers[i].lng;
        clusterSums[clusterIdx].count++;
      }

      for (let j = 0; j < k; j++) {
        if (clusterSums[j].count > 0) {
          const newLat = clusterSums[j].lat / clusterSums[j].count;
          const newLng = clusterSums[j].lng / clusterSums[j].count;
          
          // ตรวจว่าพิกัดเปลี่ยนหรือไม่
          if (Math.abs(centroids[j].lat - newLat) > 0.0001 || Math.abs(centroids[j].lng - newLng) > 0.0001) {
            centroids[j].lat = newLat;
            centroids[j].lng = newLng;
            centroidsChanged = true;
          }
        }
      }
    }

    // รวมกลุ่มลูกค้าแยกอาเรย์ย่อยตามคลัสเตอร์
    let clusters = Array.from({ length: k }, () => []);
    for (let i = 0; i < customers.length; i++) {
      clusters[assignments[i]].push(customers[i]);
    }

    return { clusters, centroids };
  }

  /**
   * ฟังก์ชันหลักในการแก้ปัญหาจัดเส้นทางเดินรถแบบแบ่งกลุ่ม (m-TSP Solver)
   */
  function solveVRP(depot, customers, config) {
    const vehicleCount = parseInt(config.vehicleCount) || 3;
    const maxStops = parseInt(config.maxStops) || 15;
    const isRoundTrip = config.isRoundTrip !== false;
    const averageSpeed = parseFloat(config.averageSpeed) || 40; // กม./ชม.
    const serviceDuration = parseFloat(config.serviceDuration) || 10; // นาที

    // คัดแยกเฉพาะพิกัดลูกค้าที่ถูกต้อง
    const validCustomers = customers.filter(c => c.lat && c.lng && !c.isDepot);
    
    if (validCustomers.length === 0) {
      return { routes: [], unassigned: [], summary: { totalDistance: 0, totalDuration: 0, vehiclesUsed: 0 } };
    }

    let optimizedRoutes = [];
    let routeNames = [];
    let unassignedCustomers = [];

    // ตรวจสอบว่าพิกัดที่โหลดมามีการระบุสายรถ/สายส่งสินค้าไว้ล่วงหน้าหรือไม่
    const hasPredefinedRoutes = validCustomers.some(c => c.route !== undefined && c.route !== null && String(c.route).trim() !== '');

    if (hasPredefinedRoutes) {
      // โหมดจัดสายส่งตามที่ระบุใน Excel
      const routeGroups = {};
      validCustomers.forEach(cust => {
        const rName = cust.route ? String(cust.route).trim() : 'ไม่ระบุสายส่ง';
        if (!routeGroups[rName]) {
          routeGroups[rName] = [];
        }
        routeGroups[rName].push(cust);
      });

      // จัดระเบียบและคิวการเดินทางของแต่ละสายส่ง
      Object.keys(routeGroups).forEach((rName, rIdx) => {
        let routeStops = routeGroups[rName];
        
        // จำกัดจำนวนจุดส่งสูงสุดต่อคัน (Max Stops)
        if (routeStops.length > maxStops) {
          // เรียงจุดจอดตามระยะห่างจากคลังสินค้าหลัก จากน้อยไปมาก
          routeStops.sort((a, b) => {
            return getDistance(a.lat, a.lng, depot.lat, depot.lng) - 
                   getDistance(b.lat, b.lng, depot.lat, depot.lng);
          });
          const excess = routeStops.splice(maxStops);
          unassignedCustomers.push(...excess);
        }

        if (routeStops.length > 0) {
          const opt = optimize2Opt(routeStops, depot, isRoundTrip);
          optimizedRoutes.push(opt);
          routeNames.push(rName);
        }
      });
    } else {
      // โหมดจัดสรรอัตโนมัติ: แบ่งกลุ่มด้วย K-Means
      const k = Math.min(vehicleCount, validCustomers.length);
      let { clusters, centroids } = runKMeans(validCustomers, k);

      // จัดการจำกัดจำนวนจุดจอดสูงสุดต่อคัน (Max Stops Constraint)
      let adjustmentNeeded = true;
      let adjustLoops = 0;
      const maxAdjustLoops = 50;

      while (adjustmentNeeded && adjustLoops < maxAdjustLoops) {
        adjustmentNeeded = false;
        adjustLoops++;

        for (let i = 0; i < clusters.length; i++) {
          if (clusters[i].length > maxStops) {
            adjustmentNeeded = true;
            
            const centroid = centroids[i];
            clusters[i].sort((a, b) => {
              return getDistance(b.lat, b.lng, centroid.lat, centroid.lng) - 
                     getDistance(a.lat, a.lng, centroid.lat, centroid.lng);
            });

            const excessCount = clusters[i].length - maxStops;
            const excessPoints = clusters[i].splice(0, excessCount);

            excessPoints.forEach(point => {
              let bestTargetClusterIdx = -1;
              let minDistToTargetCentroid = Infinity;

              for (let targetIdx = 0; targetIdx < clusters.length; targetIdx++) {
                if (targetIdx !== i && clusters[targetIdx].length < maxStops) {
                  let dist = getDistance(point.lat, point.lng, centroids[targetIdx].lat, centroids[targetIdx].lng);
                  if (dist < minDistToTargetCentroid) {
                    minDistToTargetCentroid = dist;
                    bestTargetClusterIdx = targetIdx;
                  }
                }
              }

              if (bestTargetClusterIdx !== -1) {
                clusters[bestTargetClusterIdx].push(point);
              } else {
                unassignedCustomers.push(point);
              }
            });
          }
        }
      }

      let activeRoutes = clusters.filter(c => c.length > 0);
      optimizedRoutes = activeRoutes.map(route => {
        return optimize2Opt(route, depot, isRoundTrip);
      });
      routeNames = optimizedRoutes.map((_, idx) => `สายส่งสินค้าที่ ${idx + 1}`);
    }

    // 5. คำนวณสรุปสถิติผลลัพธ์
    let totalDistance = 0;
    let totalDuration = 0;

    const routesResult = optimizedRoutes.map((route, idx) => {
      const distance = calculateRouteDistance(route, depot, isRoundTrip);
      
      const travelTimeMinutes = (distance / averageSpeed) * 60;
      const serviceTimeMinutes = route.length * serviceDuration;
      const durationMinutes = travelTimeMinutes + serviceTimeMinutes;

      totalDistance += distance;
      totalDuration += durationMinutes;

      // สร้างลำดับรายการนำทางพร้อมเวลาประมาณการเข้าถึง
      let currentDistance = 0;
      let currentTimeOffset = 0;
      
      const stops = route.map((customer, stopIdx) => {
        let prevLoc = stopIdx === 0 ? depot : route[stopIdx - 1];
        let legDistance = getDistance(prevLoc.lat, prevLoc.lng, customer.lat, customer.lng);
        currentDistance += legDistance;
        
        let legTravelTime = (legDistance / averageSpeed) * 60;
        currentTimeOffset += legTravelTime;
        
        const etaOffset = currentTimeOffset;
        currentTimeOffset += serviceDuration;

        return {
          sequence: stopIdx + 1,
          customer: customer,
          legDistance: legDistance,
          cumulativeDistance: currentDistance,
          etaOffset: etaOffset,
          departureOffset: currentTimeOffset
        };
      });

      return {
        id: idx + 1,
        vehicleName: routeNames[idx],
        color: getRandomColor(idx),
        stops: stops,
        totalDistance: distance,
        totalDuration: durationMinutes
      };
    });

    return {
      routes: routesResult,
      unassigned: unassignedCustomers,
      summary: {
        totalDistance: parseFloat(totalDistance.toFixed(2)),
        totalDuration: Math.round(totalDuration),
        vehiclesUsed: routesResult.length,
        totalCustomersCount: validCustomers.length
      }
    };
  }

  // รหัสสีพรีเมียมตามลำดับสายรถ
  function getRandomColor(index) {
    const colors = [
      '#10B981', // Emerald
      '#3B82F6', // Blue
      '#8B5CF6', // Purple
      '#EC4899', // Pink
      '#F59E0B', // Amber
      '#06B6D4', // Cyan
      '#EF4444', // Red
      '#14B8A6', // Teal
      '#F97316', // Orange
      '#6366F1'  // Indigo
    ];
    return colors[index % colors.length];
  }

  return {
    solveVRP: solveVRP,
    getDistance: getDistance
  };

})();
