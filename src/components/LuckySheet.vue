<template>
  <div class="container">
    <div class="header" :class="isMobile ? 'h60' : 'h40'">
      <div></div>
      <div style="text-align: center" @click="update">
        <button class="update-button">财务报表计算</button>
      </div>
    </div>
    <div id="luckysheet" :class="isMobile ? 't100' : 't60'"></div>
  </div>

  <Teleport to="body">
    <!-- use the modal component, pass in the prop -->
    <modal :show="showModal" @close="showModal = false">
      <template #header>
        <h2>财务报表计算结果</h2>
      </template>
      <template #body>
        <div>
          <div v-if="successInfo" class="success">您的{{ successInfo }};</div>
          <div v-if="errorInfo" class="error">您的{{ errorInfo }}</div>
        </div>
      </template>
    </modal>
  </Teleport>
</template>

<script setup>
import { ref, onMounted, onUpdated } from 'vue';
import LuckyExcel from 'luckyexcel';

import { exportExcel } from './export';
import Modal from './Modal.vue';

const templates = ['./assets/family-pc.xlsx', './assets/family-mobile.xlsx'];
const showModal = ref(false);
const isMobile = ref(true);
const successInfo = ref('');
const errorInfo = ref('');
const jsonData = ref({});
const yearlyOutcome = ref(0);

function getCell(r, c) {
  return luckysheet.getCellValue(r, 7);
}
function leftResult(col = 7, rows = [3, 4, 5]) {
  const errors = [];
  const success = [];
  const [v1, v2, v3] = rows.map((r) => getCell(r, 7));
  const [i1, i2, i3] = ['结余比率', '财务负担比', '财务自由度'].map(
    (x) => `${x}指标`
  );
  if (v1 <= 0.2) {
    errors.push(`${i1}高于参考值`);
  } else {
    success.push(`${i1}合适`);
  }
  if (v2 >= 0.4) {
    errors.push(`${i2}低于参考值`);
  } else {
    success.push(`${i2}合适`);
  }
  if (v3 <= 1.0) {
    errors.push(`${i3}低于参考值`);
  } else {
    success.push(`${i3}美好`);
  }

  return [success, errors];
}
function rightResult(col = 7, rows = [3, 4, 5]) {
  const errors = [];
  const success = [];

  const [v1, v2, v3] = rows.map((r) => getCell(r, col));
  const [i1, i2, i3] = ['负债率', '投资比率', '流动性比率'].map(
    (x) => `${x}指标`
  );

  if (v1 >= 0.7) {
    errors.push(`${i1}高于参考值`);
  } else {
    success.push(`${i1}较安全`);
  }
  if (v2 <= 0.5) {
    errors.push(`${i1}低于参考值`);
  } else {
    success.push(`${i2}合适`);
  }
  if (v3 < 3.0 || v3 > 6) {
    errors.push(`${i1}不符合参考值`);
  } else {
    success.push(`${i3}适当`);
  }
  return [success, errors];
}

function update() {
  const [ls, le] = leftResult();
  const [rs, re] = isMobile.value
    ? rightResult(7, [29, 30, 31])
    : rightResult(16);
  const result = [ls.concat(rs), le.concat(re)];

  const [successMsg, errorMsg] = result.map((x) => x.join(','));
  successInfo.value = successMsg;
  errorInfo.value = errorMsg ? `${errorMsg}, 具体提升建议，请预约咨询！` : '';
  console.log('result: ', result);

  showModal.value = true;
}

onMounted(() => {
  let isPhone = true;
  if (/Mobi|Android|iPhone/i.test(navigator.userAgent)) {
    isPhone = true;
  } else {
    isPhone = false;
  }
  isMobile.value = isPhone;
  console.log('isMobile.value: ', isMobile.value);
  luckysheet.create({
    container: 'luckysheet',
    showinfobar: false,
  });
  const template = templates[+isPhone];
  loadTable(template, 'template');
});

function loadTable(value, name) {
  if (value == '') {
    return;
  }

  LuckyExcel.transformExcelToLuckyByUrl(
    value,
    name,
    (exportJson, luckysheetfile) => {
      if (exportJson.sheets == null || exportJson.sheets.length == 0) {
        alert(
          'Failed to read the content of the excel file, currently does not support xls files!'
        );
        return;
      }
      jsonData.value = exportJson;

      window.luckysheet.destroy();

      window.luckysheet.create({
        container: 'luckysheet',
        showinfobar: false,
        data: exportJson.sheets,
        title: exportJson.info.name,
        userInfo: exportJson.info.name.creator,
      });
    }
  );
}
// onUpdated(() => {
//   if (isFirst.value && isMobile.value) {
//     setTimeout(() => {
//       console.log('set', yearlyOutcome.value);
//       luckysheet.setCellValue(25, 5, yearlyOutcome.value);
//     }, 2000);
//   }
// });

// function selectExcel(evt) {
//   const value = selected.value;
//   isFirst.value = value === types[0];
//   const name = evt.target.options[evt.target.selectedIndex].innerText;
//   loadTable(value, name);
// }
</script>

<style scoped>
.container {
  padding: 0 8px;
}
.header {
  display: flex;
  justify-content: space-between;
  align-items: center;
}
.h40 {
  height: 40px;
}
.h60 {
  height: 60px;
}
.t60 {
  top: 60px;
}
.t100 {
  top: 100px;
}
#luckysheet {
  margin: 0px;
  padding: 0px;
  position: absolute;
  width: 100%;
  left: 0px;
  bottom: 0px;
}

.update-button {
  width: 160px;
  height: 40px;
  border-width: 0px;
  border-radius: 3px;
  background: #1e90ff;
  cursor: pointer;
  outline: none;
  font-family: Microsoft YaHei;
  color: white;
  font-size: 17px;
}
.update-button:hover {
  background: #5599ff;
}

#tip {
  position: absolute;
  z-index: 1000000;
  left: 0px;
  top: 0px;
  bottom: 0px;
  right: 0px;
  background: rgba(255, 255, 255, 0.8);
  text-align: center;
  font-size: 40px;
  align-items: center;
  justify-content: center;
  display: flex;
}
.success {
  color: #1e90ff;
  margin-bottom: 16px;
}
.error {
  color: #f00;
  font-weight: 700;
}

select {
  /* styling */
  background-color: white;
  border: thin solid blue;
  border-radius: 4px;
  display: inline-block;
  font: inherit;
  line-height: 1.5em;
  padding: 0.5em 3.5em 0.5em 1em;

  /* reset */

  margin: 0;
  -webkit-box-sizing: border-box;
  -moz-box-sizing: border-box;
  box-sizing: border-box;
  -webkit-appearance: none;
  -moz-appearance: none;
}

select.minimal {
  background-image: linear-gradient(45deg, transparent 50%, gray 50%),
    linear-gradient(135deg, gray 50%, transparent 50%),
    linear-gradient(to right, #ccc, #ccc);
  background-position: calc(100% - 20px) calc(1em + 2px),
    calc(100% - 15px) calc(1em + 2px), calc(100% - 2.5em) 0.5em;
  background-size: 5px 5px, 5px 5px, 1px 1.5em;
  background-repeat: no-repeat;
}

select:-moz-focusring {
  color: transparent;
  text-shadow: 0 0 0 #000;
}
</style>
