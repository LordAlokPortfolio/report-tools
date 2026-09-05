const SPACING_Y = 2;   // vertical distance between consecutive commits
const SPACING_X = 3;   // horizontal offset per branch

fetch('history.json')
  .then((r) => r.json())
  .then((commits) => build(commits));

function build(commits) {
  const byHash = new Map(commits.map((c) => [c.hash, c]));

  // Find the tip of "main" to use as the trunk's starting point.
  const mainTip = commits.find((c) =>
    c.refs.some((r) => r === 'main' || r.endsWith('/main'))
  ) || commits[0];

  // Walk the first-parent chain from main's tip: this is the trunk.
  const trunkHashes = [];
  let cur = mainTip;
  while (cur) {
    trunkHashes.push(cur.hash);
    cur = cur.parents.length ? byHash.get(cur.parents[0]) : null;
  }
  const trunkSet = new Set(trunkHashes);

  // trunk root is the oldest commit -> y = 0, growing upward toward the tip
  const positions = new Map(); // hash -> {x, y, z}
  trunkHashes.forEach((hash, i) => {
    const y = (trunkHashes.length - 1 - i) * SPACING_Y;
    positions.set(hash, { x: 0, y, z: 0 });
  });

  // For every merge commit on the trunk, walk each non-first parent as a side branch.
  let branchIndex = 0;
  trunkHashes.forEach((hash) => {
    const c = byHash.get(hash);
    for (let p = 1; p < c.parents.length; p++) {
      branchIndex++;
      const branchHashes = [];
      let b = byHash.get(c.parents[p]);
      while (b && !positions.has(b.hash)) {
        branchHashes.push(b.hash);
        b = b.parents.length ? byHash.get(b.parents[0]) : null;
      }
      const mergeY = positions.get(hash).y;
      const x = branchIndex * SPACING_X;
      branchHashes.forEach((h, i) => {
        const y = mergeY - (i + 1) * SPACING_Y;
        positions.set(h, { x, y, z: 0 });
      });
    }
  });

  render(commits, positions);
}

function render(commits, positions) {
  const scene = new THREE.Scene();
  scene.background = new THREE.Color(0x111111);

  const camera = new THREE.PerspectiveCamera(60, innerWidth / innerHeight, 0.1, 1000);
  camera.position.set(15, 20, 25);

  const renderer = new THREE.WebGLRenderer({ antialias: true });
  renderer.setSize(innerWidth, innerHeight);
  document.body.appendChild(renderer.domElement);

  const controls = new THREE.OrbitControls(camera, renderer.domElement);
  controls.target.set(0, 20, 0);

  const nodeGeo = new THREE.SphereGeometry(0.3, 12, 12);
  const nodeMat = new THREE.MeshBasicMaterial({ color: 0x4fc3f7 });
  const mergeMat = new THREE.MeshBasicMaterial({ color: 0xffb300 });

  commits.forEach((c) => {
    const p = positions.get(c.hash);
    if (!p) return;
    const isMerge = c.parents.length > 1;
    const mesh = new THREE.Mesh(nodeGeo, isMerge ? mergeMat : nodeMat);
    mesh.position.set(p.x, p.y, p.z);
    scene.add(mesh);

    c.parents.forEach((parentHash) => {
      const pp = positions.get(parentHash);
      if (!pp) return;
      const points = [
        new THREE.Vector3(p.x, p.y, p.z),
        new THREE.Vector3(pp.x, pp.y, pp.z),
      ];
      const geo = new THREE.BufferGeometry().setFromPoints(points);
      const line = new THREE.Line(geo, new THREE.LineBasicMaterial({ color: 0x888888 }));
      scene.add(line);
    });
  });

  window.addEventListener('resize', () => {
    camera.aspect = innerWidth / innerHeight;
    camera.updateProjectionMatrix();
    renderer.setSize(innerWidth, innerHeight);
  });

  (function animate() {
    requestAnimationFrame(animate);
    controls.update();
    renderer.render(scene, camera);
  })();
}
