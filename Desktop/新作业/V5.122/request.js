// ===================================================================================
// ### request.js: 封装比特浏览器API (由用户提供) ###
// ===================================================================================
const axios = require('axios').default

const baseURL = 'http://127.0.0.1:54345'

const request = axios.create({
  baseURL,
  timeout: 0
})

request.interceptors.response.use(
  response => {
    if (response.status === 200) {
      return response.data
    } else {
      console.log('请求失败，检查网络')
    }
  },
  error => {
    console.error('请求失败了')
    return Promise.reject(error)
  }
)

function openBrowser(data) {
  return request({ method: 'post', url: '/browser/open', data })
}

function closeBrowser(id) {
  return request({ method: 'post', url: '/browser/close', data: { id } })
}

function createBrowser(data) {
  return request({ method: 'post', url: '/browser/update', data })
}

function updatepartial(data) {
  return request({ method: 'post', url: '/browser/update/partial', data })
}

function deleteBatchBrowser(ids) {
  return request({ method: 'post', url: '/browser/delete/ids', data: { ids } })
}

function deleteBrowser(id) {
  return request({ method: 'post', url: '/browser/delete', data: { id } })
}

function getBrowserDetail(id) {
  return request({ method: 'post', url: '/browser/detail', data: { id } })
}

function getBrowserList(data) {
  return request({ method: 'post', url: '/browser/list', data })
}

function getBrowserConciseList(data) {
  return request({ method: 'post', url: '/browser/list/concise', data })
}

function getGroupList(page, pageSize) {
  return request({ method: 'post', url: '/group/list', data: { page, pageSize } })
}

function addGroup(groupName, sortNum) {
  return request({ method: 'post', url: '/group/add', data: { groupName, sortNum } })
}

function editGroup(id, groupName, sortNum) {
  return request({ method: 'post', url: '/group/edit', data: { id, groupName, sortNum } })
}

function deleteGroup(id) {
  return request({ method: 'post', url: '/group/delete', data: { id } })
}

function getGroupDetail(id) {
  return request({ method: 'post', url: '/group/detail', data: { id } })
}

function getPids(ids) {
  return request({ url: '/browser/pids', method: 'post', data: { ids } })
}

function getAlivePids(ids) {
  return request({ url: '/browser/pids/alive', method: 'post', data: { ids } })
}

function getAliveBrowsersPids() {
  return request({ url: '/browser/pids/all', method: 'post' })
}

function updateBrowserMemark(remark, browserIds) {
  return request({ url: '/browser/remark/update', method: 'post', data: { remark, browserIds } })
}

function batchUpdateBrowserGroup(data) {
  return request({ url: '/browser/group/update', method: 'post', data })
}

function closeBrowsersBySeqs(seqs) {
  return request({ url: '/browser/close/byseqs', method: 'post', data: { seqs } })
}

function batchUpdateProxy(data) {
  return request({ url: '/browser/proxy/update', method: 'post', data })
}

module.exports = {
  openBrowser,
  closeBrowser,
  createBrowser,
  deleteBrowser,
  getBrowserDetail,
  addGroup,
  editGroup,
  deleteGroup,
  getGroupDetail,
  getGroupList,
  getBrowserList,
  getPids,
  updatepartial,
  updateBrowserMemark,
  deleteBatchBrowser,
  getBrowserConciseList,
  getAlivePids,
  getAliveBrowsersPids,
  batchUpdateBrowserGroup,
  closeBrowsersBySeqs,
  batchUpdateProxy,
  request
}